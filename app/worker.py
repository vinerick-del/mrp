"""
Worker - Monitoramento de posições abertas e trades fechados
Executa em loop para:
1. Gerenciar posições abertas (Break Even + Trailing Stop)
2. Verificar trades fechados
3. Atualizar performance
"""

import time
import logging
from datetime import datetime
from services.mt5_service import manage_open_positions, check_closed_trades
from services.performance_service import get_performance_summary
from services.safety_service import get_safety_status

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def worker_loop():
    """Loop principal do worker"""
    logger.info("🚀 Worker iniciado")

    check_interval = 60  # Verificar a cada 60 segundos

    while True:
        try:
            # 1. Gerenciar posições abertas
            pos_result = manage_open_positions()
            if pos_result["status"] == "sucesso" and pos_result["modified"] > 0:
                logger.info(f"📊 {pos_result['message']}")

            # 2. Verificar trades fechados
            trades_result = check_closed_trades(lookback_minutes=10)
            if trades_result["status"] == "sucesso" and trades_result["processed"] > 0:
                logger.info(f"✅ {trades_result['message']}")

            # 3. Log de status
            performance = get_performance_summary()
            safety = get_safety_status()

            logger.info(
                f"📈 Performance - Trades: {performance['total_trades']}, "
                f"Win Rate: {performance['win_rate']}%, "
                f"Net Profit: {performance['net_profit']}"
            )

            logger.info(
                f"🔒 Safety - Status: {safety['status']}, "
                f"Consecutive Losses: {safety['consecutive_losses']}"
            )

        except Exception as e:
            logger.error(f"❌ Erro no worker: {str(e)}")

        # Aguardar até a próxima verificação
        time.sleep(check_interval)


if __name__ == "__main__":
    try:
        worker_loop()
    except KeyboardInterrupt:
        logger.info("⛔ Worker interrompido pelo usuário")
    except Exception as e:
        logger.error(f"❌ Erro crítico: {str(e)}")
