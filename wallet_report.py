from __future__ import annotations

import io
import os
import time
from collections.abc import Callable
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from decimal import Decimal, getcontext
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font
from web3 import Web3
from web3.providers import HTTPProvider

getcontext().prec = 50

TRANSFER_TOPIC = Web3.to_hex(Web3.keccak(text="Transfer(address,address,uint256)"))
RPC_URL_DEFAULT = "https://rpc1.netsbo.io"

ProgressCallback = Callable[[str, float, str, dict[str, Any] | None], None]
CancelCallback = Callable[[], None]

TOKEN_INFO_ABI = [
    {
        "name": "name",
        "outputs": [{"type": "string"}],
        "inputs": [],
        "stateMutability": "view",
        "type": "function",
    },
    {
        "name": "symbol",
        "outputs": [{"type": "string"}],
        "inputs": [],
        "stateMutability": "view",
        "type": "function",
    },
    {
        "name": "decimals",
        "outputs": [{"type": "uint8"}],
        "inputs": [],
        "stateMutability": "view",
        "type": "function",
    },
]


@dataclass(slots=True)
class WalletTransfer:
    block_number: int
    log_index: int
    tx_hash: str
    contract_address: str
    token_name: str
    token_symbol: str
    decimals: int
    direction: str
    amount: str
    raw_value: str
    from_address: str
    to_address: str
    timestamp: str


def _noop_progress(_phase: str, _progress: float, _message: str, _metrics: dict[str, Any] | None):
    return


def _noop_cancel():
    return


def get_web3_connection(rpc_url: str | None = None) -> Web3:
    provider = HTTPProvider(
        rpc_url or os.environ.get("NETSBO_RPC_URL", RPC_URL_DEFAULT),
        request_kwargs={"timeout": 120},
    )
    web3 = Web3(provider)
    try:
        from web3.middleware import ExtraDataToPOAMiddleware

        web3.middleware_onion.inject(ExtraDataToPOAMiddleware, layer=0)
    except Exception:
        try:
            from web3.middleware import geth_poa_middleware

            web3.middleware_onion.inject(geth_poa_middleware, layer=0)
        except Exception:
            pass
    if not web3.is_connected():
        raise ConnectionError("No se pudo conectar al RPC de Netsbo.")
    return web3


def emit_progress(
    callback: ProgressCallback,
    phase: str,
    progress: float,
    message: str,
    metrics: dict[str, Any] | None = None,
):
    callback(phase, max(0.0, min(progress, 100.0)), message, metrics or {})


def to_plain_decimal(raw_value: int, decimals: int) -> str:
    value = Decimal(raw_value) / (Decimal(10) ** decimals)
    rendered = format(value.normalize(), "f")
    if "." in rendered:
        rendered = rendered.rstrip("0").rstrip(".")
    return rendered or "0"


def topic_address(wallet_address: str) -> str:
    return "0x" + wallet_address.lower()[2:].zfill(64)


def topic_to_checksum_address(value: Any) -> str:
    if hasattr(value, "hex"):
        raw = value.hex()
    else:
        raw = str(value)
    if raw.startswith("0x"):
        raw = raw[2:]
    return Web3.to_checksum_address("0x" + raw[-40:])


def hex_to_int(value: Any) -> int:
    if value is None:
        return 0
    if isinstance(value, int):
        return value
    if isinstance(value, bytes):
        return int.from_bytes(value, byteorder="big", signed=False)
    text = value.hex() if hasattr(value, "hex") else str(value)
    if text.startswith("0x"):
        text = text[2:]
    return int(text or "0", 16)


def find_block_by_date(
    web3: Web3,
    target_date: datetime,
    progress_callback: ProgressCallback,
    cancel_callback: CancelCallback,
) -> int:
    latest_block = int(web3.eth.block_number)
    low = 0
    high = latest_block
    result = latest_block
    estimated_steps = max(latest_block.bit_length(), 1)
    step = 0

    while low <= high:
        cancel_callback()
        mid = (low + high) // 2
        block = web3.eth.get_block(mid)
        block_time = datetime.fromtimestamp(block.timestamp, tz=timezone.utc)
        step += 1
        emit_progress(
            progress_callback,
            "searching",
            2 + (step / estimated_steps) * 8,
            f"Buscando bloque inicial por fecha. Iteración {step}/{estimated_steps}.",
            {"endBlock": latest_block},
        )
        if block_time < target_date:
            low = mid + 1
        else:
            result = mid
            high = mid - 1

    return result


def get_logs_with_retry(
    web3: Web3,
    filter_params: dict[str, Any],
    cancel_callback: CancelCallback,
    max_retries: int = 5,
    min_batch_size: int = 100,
) -> list[Any]:
    from_block = int(filter_params["fromBlock"])
    to_block = int(filter_params["toBlock"])
    attempt = 0

    while True:
        cancel_callback()
        try:
            return list(web3.eth.get_logs(filter_params))
        except Exception as exc:
            attempt += 1
            span = to_block - from_block + 1
            if span > min_batch_size:
                middle = from_block + (span // 2) - 1
                left_params = dict(filter_params, fromBlock=from_block, toBlock=middle)
                right_params = dict(filter_params, fromBlock=middle + 1, toBlock=to_block)
                left = get_logs_with_retry(
                    web3,
                    left_params,
                    cancel_callback,
                    max_retries=max_retries,
                    min_batch_size=min_batch_size,
                )
                right = get_logs_with_retry(
                    web3,
                    right_params,
                    cancel_callback,
                    max_retries=max_retries,
                    min_batch_size=min_batch_size,
                )
                return left + right
            if attempt <= max_retries:
                time.sleep(min(8.0, 0.7 * attempt + (attempt ** 1.2) * 0.15))
                continue
            raise RuntimeError(
                f"Fallo persistente al consultar logs entre bloques {from_block:,}-{to_block:,}: {exc}"
            ) from exc


def get_token_info(
    web3: Web3,
    contract_address: str,
    token_cache: dict[str, tuple[str, str, int]],
) -> tuple[str, str, int]:
    if contract_address in token_cache:
        return token_cache[contract_address]

    try:
        contract = web3.eth.contract(address=contract_address, abi=TOKEN_INFO_ABI)
        name = str(contract.functions.name().call())
        symbol = str(contract.functions.symbol().call())
        decimals = int(contract.functions.decimals().call())
    except Exception:
        name = contract_address
        symbol = "UNKNOWN"
        decimals = 18

    token_cache[contract_address] = (name, symbol, decimals)
    return token_cache[contract_address]


def get_block_timestamp(
    web3: Web3,
    block_number: int,
    block_cache: dict[int, str],
) -> str:
    cached = block_cache.get(block_number)
    if cached is not None:
        return cached
    block = web3.eth.get_block(block_number)
    timestamp = datetime.fromtimestamp(block.timestamp, tz=timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    block_cache[block_number] = timestamp
    return timestamp


def build_workbook(
    wallet: str,
    days: int,
    start_block: int,
    end_block: int,
    transfers: list[WalletTransfer],
    transfers_in: int,
    transfers_out: int,
    transfers_self: int,
    unique_tokens: int,
) -> bytes:
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Resumen"
    summary.append(["Campo", "Valor"])
    summary["A1"].font = Font(bold=True)
    summary["B1"].font = Font(bold=True)

    generated_at = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    summary_rows = [
        ("Wallet", wallet),
        ("Dias analizados", days),
        ("Bloque inicial", start_block),
        ("Bloque final", end_block),
        ("Transfers totales", len(transfers)),
        ("Entradas", transfers_in),
        ("Salidas", transfers_out),
        ("Self", transfers_self),
        ("Tokens unicos", unique_tokens),
        ("Generado", generated_at),
    ]
    for row in summary_rows:
        summary.append(list(row))
    summary.column_dimensions["A"].width = 24
    summary.column_dimensions["B"].width = 42

    sheet = workbook.create_sheet("Transfers")
    headers = [
        "Block",
        "TxHash",
        "Contract",
        "Token",
        "Symbol",
        "Direction",
        "Amount",
        "Raw Value",
        "Decimals",
        "From",
        "To",
        "Date",
    ]
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)

    for transfer in transfers:
        sheet.append(
            [
                transfer.block_number,
                transfer.tx_hash,
                transfer.contract_address,
                transfer.token_name,
                transfer.token_symbol,
                transfer.direction,
                transfer.amount,
                transfer.raw_value,
                transfer.decimals,
                transfer.from_address,
                transfer.to_address,
                transfer.timestamp,
            ]
        )

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:L{max(len(transfers) + 1, 2)}"
    widths = {
        "A": 12,
        "B": 68,
        "C": 44,
        "D": 28,
        "E": 16,
        "F": 12,
        "G": 24,
        "H": 28,
        "I": 12,
        "J": 44,
        "K": 44,
        "L": 24,
    }
    for column, width in widths.items():
        sheet.column_dimensions[column].width = width

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.read()


def generate_wallet_report(
    wallet: str,
    days: int,
    batch_size: int = 3000,
    min_batch_size: int = 100,
    rpc_url: str | None = None,
    progress_callback: ProgressCallback | None = None,
    cancel_callback: CancelCallback | None = None,
) -> dict[str, Any]:
    progress = progress_callback or _noop_progress
    cancel = cancel_callback or _noop_cancel

    if days < 1:
        raise ValueError("days debe ser mayor o igual a 1")

    try:
        wallet_checksum = Web3.to_checksum_address(wallet)
    except Exception as exc:
        raise ValueError("Wallet inválida") from exc

    emit_progress(progress, "starting", 1, "Conectando al RPC de Netsbo...")
    web3 = get_web3_connection(rpc_url)
    wallet_lower = wallet_checksum.lower()
    wallet_topic = topic_address(wallet_checksum)

    target_date = datetime.now(timezone.utc) - timedelta(days=days)
    start_block = find_block_by_date(web3, target_date, progress, cancel)
    end_block = int(web3.eth.block_number)
    total_blocks = max(end_block - start_block + 1, 0)

    emit_progress(
        progress,
        "fetching_logs",
        12,
        f"Preparando escaneo entre los bloques {start_block:,} y {end_block:,}.",
        {"startBlock": start_block, "endBlock": end_block},
    )

    all_transfers: dict[tuple[str, int], tuple[Any, str]] = {}
    total_batches = max(((end_block - start_block) // batch_size) + 1, 1)

    for batch_index, batch_start in enumerate(range(start_block, end_block + 1, batch_size), start=1):
        cancel()
        batch_end = min(batch_start + batch_size - 1, end_block)
        batch_progress = 12 + (batch_index / total_batches) * 48

        emit_progress(
            progress,
            "fetching_logs",
            batch_progress,
            f"Consultando logs entre bloques {batch_start:,}-{batch_end:,} ({batch_index}/{total_batches}).",
            {
                "startBlock": start_block,
                "endBlock": end_block,
                "batchesDone": batch_index,
                "batchesTotal": total_batches,
                "totalTransfers": len(all_transfers),
            },
        )

        incoming_logs = get_logs_with_retry(
            web3,
            {
                "fromBlock": batch_start,
                "toBlock": batch_end,
                "topics": [TRANSFER_TOPIC, None, wallet_topic],
            },
            cancel,
            min_batch_size=min_batch_size,
        )
        for log in incoming_logs:
            key = (Web3.to_hex(log["transactionHash"]), int(log.get("logIndex", 0)))
            if key not in all_transfers:
                all_transfers[key] = (log, "IN")

        outgoing_logs = get_logs_with_retry(
            web3,
            {
                "fromBlock": batch_start,
                "toBlock": batch_end,
                "topics": [TRANSFER_TOPIC, wallet_topic, None],
            },
            cancel,
            min_batch_size=min_batch_size,
        )
        for log in outgoing_logs:
            key = (Web3.to_hex(log["transactionHash"]), int(log.get("logIndex", 0)))
            if key in all_transfers:
                all_transfers[key] = (log, "SELF")
            else:
                all_transfers[key] = (log, "OUT")

    transfers_by_order = sorted(
        all_transfers.values(),
        key=lambda item: (int(item[0]["blockNumber"]), int(item[0].get("logIndex", 0))),
    )

    emit_progress(
        progress,
        "processing",
        62,
        f"Procesando {len(transfers_by_order):,} transferencias encontradas.",
        {
            "startBlock": start_block,
            "endBlock": end_block,
            "totalTransfers": len(transfers_by_order),
        },
    )

    token_cache: dict[str, tuple[str, str, int]] = {}
    block_cache: dict[int, str] = {}
    rows: list[WalletTransfer] = []
    unique_tokens: set[str] = set()
    transfers_in = 0
    transfers_out = 0
    transfers_self = 0
    total_transfers = len(transfers_by_order)

    for index, (log, saved_direction) in enumerate(transfers_by_order, start=1):
        cancel()
        if index % 25 == 0 or index == total_transfers:
            processing_progress = 62 + (index / max(total_transfers, 1)) * 30
            emit_progress(
                progress,
                "processing",
                processing_progress,
                f"Construyendo reporte {index:,}/{total_transfers:,}.",
                {
                    "startBlock": start_block,
                    "endBlock": end_block,
                    "totalTransfers": total_transfers,
                    "transfersIn": transfers_in,
                    "transfersOut": transfers_out,
                    "transfersSelf": transfers_self,
                    "uniqueTokens": len(unique_tokens),
                },
            )

        block_number = int(log["blockNumber"])
        log_index = int(log.get("logIndex", 0))
        tx_hash = Web3.to_hex(log["transactionHash"])
        contract_address = Web3.to_checksum_address(log["address"])
        topics = log.get("topics", [])
        if len(topics) < 3:
            continue

        from_address = topic_to_checksum_address(topics[1])
        to_address = topic_to_checksum_address(topics[2])
        raw_value_int = hex_to_int(log.get("data"))
        name, symbol, decimals = get_token_info(web3, contract_address, token_cache)
        amount = to_plain_decimal(raw_value_int, decimals)
        timestamp = get_block_timestamp(web3, block_number, block_cache)

        direction = saved_direction
        from_lower = from_address.lower()
        to_lower = to_address.lower()
        if from_lower == wallet_lower and to_lower == wallet_lower:
            direction = "SELF"
        elif to_lower == wallet_lower:
            direction = "IN"
        elif from_lower == wallet_lower:
            direction = "OUT"

        if direction == "IN":
            transfers_in += 1
        elif direction == "OUT":
            transfers_out += 1
        else:
            transfers_self += 1

        unique_tokens.add(contract_address)
        rows.append(
            WalletTransfer(
                block_number=block_number,
                log_index=log_index,
                tx_hash=tx_hash,
                contract_address=contract_address,
                token_name=name,
                token_symbol=symbol,
                decimals=decimals,
                direction=direction,
                amount=amount,
                raw_value=str(raw_value_int),
                from_address=from_address,
                to_address=to_address,
                timestamp=timestamp,
            )
        )

    emit_progress(
        progress,
        "generating",
        95,
        "Generando archivo Excel del reporte...",
        {
            "startBlock": start_block,
            "endBlock": end_block,
            "totalTransfers": len(rows),
            "transfersIn": transfers_in,
            "transfersOut": transfers_out,
            "transfersSelf": transfers_self,
            "uniqueTokens": len(unique_tokens),
        },
    )

    report_bytes = build_workbook(
        wallet=wallet_checksum,
        days=days,
        start_block=start_block,
        end_block=end_block,
        transfers=rows,
        transfers_in=transfers_in,
        transfers_out=transfers_out,
        transfers_self=transfers_self,
        unique_tokens=len(unique_tokens),
    )

    generated_at = datetime.now(timezone.utc)
    filename = f"wallet-report-{wallet_checksum[-6:].lower()}-{generated_at.strftime('%Y%m%d-%H%M%S')}.xlsx"

    emit_progress(
        progress,
        "complete",
        100,
        f"Reporte completo: {transfers_in} IN, {transfers_out} OUT, {transfers_self} SELF.",
        {
            "startBlock": start_block,
            "endBlock": end_block,
            "totalTransfers": len(rows),
            "transfersIn": transfers_in,
            "transfersOut": transfers_out,
            "transfersSelf": transfers_self,
            "uniqueTokens": len(unique_tokens),
        },
    )

    return {
        "wallet": wallet_checksum,
        "days": days,
        "startBlock": start_block,
        "endBlock": end_block,
        "totalBlocks": total_blocks,
        "totalTransfers": len(rows),
        "transfersIn": transfers_in,
        "transfersOut": transfers_out,
        "transfersSelf": transfers_self,
        "uniqueTokens": len(unique_tokens),
        "filename": filename,
        "xlsxBytes": report_bytes,
    }
