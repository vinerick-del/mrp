"""
Serviço de Integração com Telegram
"""

import requests
from typing import Dict, Any
from config import TELEGRAM_TOKEN, TELEGRAM_CHAT_ID


def send_telegram_message(message: str) -> Dict[str, Any]:
    """
    Enviar mensagem via Telegram Bot

    Args:
        message: Texto da mensagem a enviar

    Returns:
        Dicionário com status e resultado da operação
    """

    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"

    payload = {
        "chat_id": TELEGRAM_CHAT_ID,
        "text": message,
        "parse_mode": "HTML"
    }

    try:
        response = requests.post(url, json=payload, timeout=10)
        response.raise_for_status()

        return {
            "status": "sucesso",
            "message": "Mensagem enviada com sucesso",
            "code": response.status_code
        }

    except requests.exceptions.RequestException as e:
        return {
            "status": "erro",
            "message": f"Erro ao enviar mensagem: {str(e)}",
            "code": None
        }

    except Exception as e:
        return {
            "status": "erro",
            "message": f"Erro inesperado: {str(e)}",
            "code": None
        }
