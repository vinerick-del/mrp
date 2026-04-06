"""
Serviço de Integração com MetaTrader 5
Execução automática de ordens com cálculo de risco
"""

from typing import Dict, Any, Optional, Set
from datetime import datetime, timedelta
import MetaTrader5 as mt5
from config import RISK_PER_TRADE
from services.performance_service import log_trade, get_performance_summary, get_symbol_ranking

# Configurações de ordem
DEFAULT_VOLUME = 0.01  # Lote padrão (mínimo)
MAX_VOLUME = 0.05  # Lote máximo
SL_PIPS = 20  # Stop Loss em pips
TP_PIPS = 40  # Take Profit em pips (RR 1:2)

# Rastreamento de trades já processados (evitar duplicidade)
processed_deal_ids: Set[int] = set()


def connect_mt5(server: str = "MetaQuotes-Demo", login: int = None, password: str = None) -> Dict[str, Any]:
    """
    Conectar ao MetaTrader 5

    Args:
        server: Nome do servidor MT5
        login: Login da conta (opcional)
        password: Senha da conta (opcional)

    Returns:
        Status de conexão
    """
    try:
        if not mt5.initialize():
            return {
                "status": "erro",
                "message": f"Falha ao inicializar MT5: {mt5.last_error()}",
                "connected": False
            }

        # Se login e password fornecidos, conectar à conta específica
        if login and password:
            if not mt5.login(login, password, server):
                return {
                    "status": "erro",
                    "message": f"Falha ao login: {mt5.last_error()}",
                    "connected": False
                }

        return {
            "status": "sucesso",
            "message": "Conectado ao MetaTrader 5",
            "connected": True
        }

    except Exception as e:
        return {
            "status": "erro",
            "message": f"Erro de conexão: {str(e)}",
            "connected": False
        }


def disconnect_mt5() -> Dict[str, Any]:
    """Desconectar do MetaTrader 5"""
    try:
        mt5.shutdown()
        return {
            "status": "sucesso",
            "message": "Desconectado do MetaTrader 5"
        }
    except Exception as e:
        return {
            "status": "erro",
            "message": f"Erro ao desconectar: {str(e)}"
        }


def get_symbol_info(symbol: str) -> Optional[Dict[str, Any]]:
    """
    Obter informações do símbolo (preço, pips)

    Args:
        symbol: Par de moedas (ex: EURUSD)

    Returns:
        Dicionário com preços ou None se erro
    """
    try:
        symbol_info = mt5.symbol_info(symbol)
        if symbol_info is None:
            return None

        return {
            "symbol": symbol,
            "bid": symbol_info.bid,
            "ask": symbol_info.ask,
            "point": symbol_info.point
        }

    except Exception as e:
        print(f"Erro ao obter info: {str(e)}")
        return None


def has_open_position(symbol: str) -> bool:
    """
    Verificar se existe posição aberta para o símbolo

    Args:
        symbol: Par de moedas (ex: EURUSD)

    Returns:
        True se existe posição aberta, False caso contrário
    """
    try:
        positions = mt5.positions_get(symbol=symbol)
        if positions is None:
            return False
        return len(positions) > 0
    except Exception as e:
        print(f"Erro ao verificar posições: {str(e)}")
        return False


def calculate_sl_tp(signal_type: str, current_price: float, point: float) -> Dict[str, float]:
    """
    Calcular Stop Loss e Take Profit

    Args:
        signal_type: "BUY" ou "SELL"
        current_price: Preço atual
        point: Um pip da moeda

    Returns:
        Dicionário com SL e TP
    """
    pip_value = point * SL_PIPS

    if signal_type == "BUY":
        sl = current_price - pip_value
        tp = current_price + (pip_value * 2)  # RR 1:2
    else:  # SELL
        sl = current_price + pip_value
        tp = current_price - (pip_value * 2)  # RR 1:2

    return {
        "sl": round(sl, 5),
        "tp": round(tp, 5)
    }


def calculate_position_size(sl_pips: float, symbol: str) -> Optional[float]:
    """
    Calcular tamanho da posição baseado no risco por trade

    Args:
        sl_pips: Stop Loss em pips
        symbol: Par de moedas

    Returns:
        Tamanho do lote calculado ou None se erro
    """
    try:
        # Obter saldo da conta
        account_info = mt5.account_info()
        if account_info is None:
            print("Erro ao obter informações da conta")
            return None

        balance = account_info.balance

        # Calcular risco em reais
        risk_amount = balance * (RISK_PER_TRADE / 100)

        # Obter informações do símbolo
        symbol_info = get_symbol_info(symbol)
        if symbol_info is None:
            return None

        # Valor por pip
        point = symbol_info["point"]
        bid = symbol_info["bid"]

        # Para Forex, valor por pip = point * 10000
        # Ajustar conforme necessário para outros ativos
        value_per_pip = point * 10000 * bid

        # Calcular lote
        if value_per_pip <= 0:
            return DEFAULT_VOLUME

        lot = risk_amount / (sl_pips * value_per_pip)

        # Respeitar lote mínimo
        lot = max(lot, DEFAULT_VOLUME)

        # Arredondar para múltiplos de 0.01
        lot = round(lot, 2)

        return lot

    except Exception as e:
        print(f"Erro ao calcular tamanho da posição: {str(e)}")
        return DEFAULT_VOLUME


def calculate_dynamic_lot(performance: Dict[str, Any], symbol: str = None) -> float:
    """
    Calcular tamanho do lote dinamicamente baseado em performance e ranking

    Args:
        performance: Dicionário com estatísticas de performance
        symbol: Símbolo atual (opcional, para ranking)

    Returns:
        Tamanho do lote ajustado
    """
    try:
        base_lot = DEFAULT_VOLUME  # 0.01
        win_rate = performance.get("win_rate", 50.0)
        drawdown = performance.get("drawdown", 0.0)

        # Ajuste baseado em win_rate
        if win_rate > 60.0:
            # Win rate alto: aumentar lote em +50%
            adjusted_lot = base_lot * 1.5  # 0.015
        elif 40.0 <= win_rate <= 60.0:
            # Win rate normal: manter lote base
            adjusted_lot = base_lot  # 0.01
        else:  # win_rate < 40.0
            # Win rate baixo: reduzir lote em -50%
            adjusted_lot = base_lot * 0.5  # 0.005

        # Proteção por drawdown
        if drawdown >= 10.0:
            # Drawdown crítico: lote mínimo
            adjusted_lot = base_lot  # 0.01
        elif drawdown >= 5.0:
            # Drawdown significativo: reduzir pela metade
            adjusted_lot = adjusted_lot * 0.5

        # Integração com ranking de ativos (auto-alocação)
        if symbol:
            ranking = get_symbol_ranking()
            if ranking:
                # Encontrar posição do símbolo no ranking
                symbol_position = None
                for idx, asset in enumerate(ranking):
                    if asset["symbol"].upper() == symbol.upper():
                        symbol_position = idx
                        break

                if symbol_position is not None:
                    total_symbols = len(ranking)

                    # TOP 1: aumentar lote +100%
                    if symbol_position == 0:
                        adjusted_lot = adjusted_lot * 2.0
                    # TOP 2: aumentar lote +50%
                    elif symbol_position == 1:
                        adjusted_lot = adjusted_lot * 1.5
                    # ÚLTIMO: reduzir lote -50%
                    elif symbol_position == total_symbols - 1:
                        adjusted_lot = adjusted_lot * 0.5

        # Garantir lote dentro dos limites
        adjusted_lot = max(adjusted_lot, DEFAULT_VOLUME)
        adjusted_lot = min(adjusted_lot, MAX_VOLUME)

        # Arredondar para múltiplos de 0.01
        adjusted_lot = round(adjusted_lot, 2)

        return adjusted_lot

    except Exception as e:
        print(f"Erro ao calcular lote dinâmico: {str(e)}")
        return DEFAULT_VOLUME


def send_order(signal: Dict[str, Any]) -> Dict[str, Any]:
    """
    Enviar ordem ao MetaTrader 5

    Args:
        signal: Dicionário com pair, signal (BUY/SELL)

    Returns:
        Status da operação
    """
    try:
        symbol = signal.get("pair")
        signal_type = signal.get("signal")

        if not symbol or not signal_type:
            return {
                "status": "erro",
                "message": "Sinal inválido (faltam pair ou signal)"
            }

        # Verificar se já existe posição aberta
        if has_open_position(symbol):
            return {
                "status": "erro",
                "message": f"Já existe posição aberta para {symbol}"
            }

        # Obter informações do símbolo
        symbol_info = get_symbol_info(symbol)
        if symbol_info is None:
            return {
                "status": "erro",
                "message": f"Símbolo {symbol} não encontrado"
            }

        # Preço para operação
        price = symbol_info["ask"] if signal_type == "BUY" else symbol_info["bid"]

        # Calcular SL e TP
        sl_tp = calculate_sl_tp(signal_type, price, symbol_info["point"])

        # Calcular tamanho da posição dinâmico baseado em performance e ranking
        performance = get_performance_summary()
        position_size = calculate_dynamic_lot(performance, symbol)

        # Tipo de ordem
        order_type = mt5.ORDER_TYPE_BUY if signal_type == "BUY" else mt5.ORDER_TYPE_SELL

        # Montar requisição de ordem
        request = {
            "action": mt5.TRADE_ACTION_DEAL,
            "symbol": symbol,
            "volume": position_size,
            "type": order_type,
            "price": price,
            "sl": sl_tp["sl"],
            "tp": sl_tp["tp"],
            "deviation": 10,
            "magic": 234000,
            "comment": f"Trading System - {signal_type}",
            "type_time": mt5.ORDER_TIME_GTC,
            "type_filling": mt5.ORDER_FILLING_IOC
        }

        # Enviar ordem
        result = mt5.order_send(request)

        if result.retcode != mt5.TRADE_RETCODE_DONE:
            return {
                "status": "erro",
                "message": f"Falha ao enviar ordem: {result.comment}",
                "retcode": result.retcode
            }

        return {
            "status": "sucesso",
            "message": f"Ordem {signal_type} enviada",
            "order_id": result.order,
            "price": price,
            "sl": sl_tp["sl"],
            "tp": sl_tp["tp"],
            "volume": position_size
        }

    except Exception as e:
        return {
            "status": "erro",
            "message": f"Erro ao enviar ordem: {str(e)}"
        }


def manage_open_positions() -> Dict[str, Any]:
    """
    Gerenciar posições abertas com Break Even e Trailing Stop

    Lógica:
    - Break Even: Se lucro >= 1R, mover SL para entrada
    - Trailing Stop: Se lucro >= 2R, atualizar SL com 1R de distância

    Returns:
        Dicionário com resumo das alterações
    """
    try:
        # Obter todas as posições abertas
        positions = mt5.positions_get()
        if positions is None or len(positions) == 0:
            return {
                "status": "sucesso",
                "message": "Nenhuma posição aberta",
                "modified": 0
            }

        modified_count = 0
        details = []

        for position in positions:
            symbol = position.symbol
            ticket = position.ticket
            position_type = position.type
            open_price = position.price_open
            current_sl = position.sl
            current_tp = position.tp

            # Obter preço atual
            symbol_info = get_symbol_info(symbol)
            if symbol_info is None:
                continue

            current_price = symbol_info["bid"] if position_type == mt5.ORDER_TYPE_BUY else symbol_info["ask"]
            point = symbol_info["point"]

            # Calcular lucro em pips
            if position_type == mt5.ORDER_TYPE_BUY:
                profit_pips = (current_price - open_price) / point
                one_r_pips = SL_PIPS
            else:  # SELL
                profit_pips = (open_price - current_price) / point
                one_r_pips = SL_PIPS

            # Break Even: Se lucro >= 1R, mover SL para entrada
            if profit_pips >= one_r_pips:
                new_sl = open_price

                # Validar: não mover SL para pior posição
                if position_type == mt5.ORDER_TYPE_BUY:
                    # Para BUY, novo SL deve ser >= SL anterior
                    if new_sl > current_sl:
                        # Modificar posição
                        request = {
                            "action": mt5.TRADE_ACTION_SLTP,
                            "position": ticket,
                            "sl": round(new_sl, 5),
                            "tp": current_tp
                        }
                        result = mt5.order_modify(request)
                        if result:
                            modified_count += 1
                            details.append({
                                "ticket": ticket,
                                "symbol": symbol,
                                "type": "Break Even",
                                "new_sl": round(new_sl, 5),
                                "profit_pips": round(profit_pips, 2)
                            })

                else:  # SELL
                    # Para SELL, novo SL deve ser <= SL anterior
                    if new_sl < current_sl:
                        # Modificar posição
                        request = {
                            "action": mt5.TRADE_ACTION_SLTP,
                            "position": ticket,
                            "sl": round(new_sl, 5),
                            "tp": current_tp
                        }
                        result = mt5.order_modify(request)
                        if result:
                            modified_count += 1
                            details.append({
                                "ticket": ticket,
                                "symbol": symbol,
                                "type": "Break Even",
                                "new_sl": round(new_sl, 5),
                                "profit_pips": round(profit_pips, 2)
                            })

            # Trailing Stop: Se lucro >= 2R, atualizar SL com 1R de distância
            elif profit_pips >= (one_r_pips * 2):
                # Manter 1R de distância do preço atual
                trailing_distance = (one_r_pips * point)

                if position_type == mt5.ORDER_TYPE_BUY:
                    new_sl = current_price - trailing_distance
                    # Novo SL deve ser >= SL anterior
                    if new_sl > current_sl:
                        request = {
                            "action": mt5.TRADE_ACTION_SLTP,
                            "position": ticket,
                            "sl": round(new_sl, 5),
                            "tp": current_tp
                        }
                        result = mt5.order_modify(request)
                        if result:
                            modified_count += 1
                            details.append({
                                "ticket": ticket,
                                "symbol": symbol,
                                "type": "Trailing Stop",
                                "new_sl": round(new_sl, 5),
                                "profit_pips": round(profit_pips, 2)
                            })

                else:  # SELL
                    new_sl = current_price + trailing_distance
                    # Novo SL deve ser <= SL anterior
                    if new_sl < current_sl:
                        request = {
                            "action": mt5.TRADE_ACTION_SLTP,
                            "position": ticket,
                            "sl": round(new_sl, 5),
                            "tp": current_tp
                        }
                        result = mt5.order_modify(request)
                        if result:
                            modified_count += 1
                            details.append({
                                "ticket": ticket,
                                "symbol": symbol,
                                "type": "Trailing Stop",
                                "new_sl": round(new_sl, 5),
                                "profit_pips": round(profit_pips, 2)
                            })

        return {
            "status": "sucesso",
            "message": f"{modified_count} posição(ões) atualizada(s)",
            "modified": modified_count,
            "details": details
        }

    except Exception as e:
        return {
            "status": "erro",
            "message": f"Erro ao gerenciar posições: {str(e)}",
            "modified": 0
        }


def check_closed_trades(lookback_minutes: int = 60) -> Dict[str, Any]:
    """
    Verificar trades fechados e registrar no performance_service

    Args:
        lookback_minutes: Buscar trades dos últimos N minutos

    Returns:
        Resumo dos trades processados
    """
    global processed_deal_ids

    try:
        # Obter timeframe para busca
        utc_from = datetime.utcnow() - timedelta(minutes=lookback_minutes)

        # Buscar operações fechadas
        deals = mt5.history_deals_get(utc_from)
        if deals is None or len(deals) == 0:
            return {
                "status": "sucesso",
                "message": "Nenhum trade fechado encontrado",
                "processed": 0
            }

        processed_count = 0
        details = []

        for deal in deals:
            deal_id = deal.ticket

            # Evitar duplicidade
            if deal_id in processed_deal_ids:
                continue

            # Filtrar apenas operações de fechamento (tipo DEAL_TYPE_SELL ou EXIT)
            if deal.entry != mt5.DEAL_ENTRY_OUT:
                continue

            # Extrair dados do trade
            symbol = deal.symbol
            profit = deal.profit
            volume = deal.volume
            close_time = datetime.fromtimestamp(deal.time).isoformat()

            # Determinar tipo (BUY ou SELL) - baseado no histórico
            # Para simplificar, usamos informações disponíveis
            trade_type = "BUY" if deal.type == mt5.ORDER_TYPE_BUY else "SELL"

            # Classificar resultado
            result = "WIN" if profit > 0 else ("LOSS" if profit < 0 else "BREAK_EVEN")

            # Montar dados para log
            trade_log = {
                "par": symbol,
                "tipo": trade_type,
                "volume": volume,
                "entry_price": deal.price_open if hasattr(deal, 'price_open') else deal.price,
                "exit_price": deal.price,
                "profit_loss": profit,
                "result": result,
                "duration": None  # Será calculado se necessário
            }

            # Registrar no performance_service
            log_result = log_trade(trade_log)

            if log_result["status"] == "sucesso":
                processed_deal_ids.add(deal_id)
                processed_count += 1
                details.append({
                    "deal_id": deal_id,
                    "symbol": symbol,
                    "profit": round(profit, 2),
                    "result": result
                })

        return {
            "status": "sucesso",
            "message": f"{processed_count} trade(s) registrado(s)",
            "processed": processed_count,
            "details": details
        }

    except Exception as e:
        return {
            "status": "erro",
            "message": f"Erro ao verificar trades fechados: {str(e)}",
            "processed": 0
        }
