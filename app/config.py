"""
Sistema de Trading Semi-Automático - Configurações
"""

import os
from dotenv import load_dotenv

load_dotenv()

# TELEGRAM
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "seu_token_aqui")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "seu_chat_id_aqui")

# WEBHOOK
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET", "seu_secret_aqui")

# RISK MANAGEMENT
RISK_PER_TRADE = float(os.getenv("RISK_PER_TRADE", "1.0"))  # Percentual do capital por trade
MAX_TRADES_PER_DAY = int(os.getenv("MAX_TRADES_PER_DAY", "5"))

# ENVIRONMENT
DEBUG = os.getenv("DEBUG", "False").lower() == "true"
