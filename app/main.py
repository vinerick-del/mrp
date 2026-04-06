"""
Sistema de Trading Semi-Automático - API Principal
"""

from fastapi import FastAPI
from fastapi.responses import JSONResponse
from config import DEBUG
from services.performance_service import get_performance_summary
from services.mt5_service import connect_mt5
import MetaTrader5 as mt5

# Inicializar aplicação
app = FastAPI(
    title="Trading System API",
    description="Sistema de Trading Semi-Automático",
    version="1.0.0",
    debug=DEBUG
)


# ROTAS
@app.get("/")
async def health_check():
    """Verificar status da API"""
    return JSONResponse(
        status_code=200,
        content={"status": "API online", "service": "Trading System"}
    )


@app.get("/health")
async def system_health():
    """Verificar saúde do sistema"""
    try:
        # Verificar conexão MT5
        account_info = mt5.account_info()
        mt5_connected = account_info is not None

        return JSONResponse(
            status_code=200,
            content={
                "status": "online",
                "api": "healthy",
                "mt5": "connected" if mt5_connected else "disconnected",
                "timestamp": __import__('datetime').datetime.utcnow().isoformat()
            }
        )
    except Exception as e:
        return JSONResponse(
            status_code=200,
            content={
                "status": "online",
                "api": "healthy",
                "mt5": "error",
                "error": str(e)
            }
        )


@app.get("/performance")
async def performance_stats():
    """Obter estatísticas de performance do bot"""
    try:
        summary = get_performance_summary()

        return JSONResponse(
            status_code=200,
            content={
                "status": "sucesso",
                "performance": summary,
                "timestamp": __import__('datetime').datetime.utcnow().isoformat()
            }
        )
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={
                "status": "erro",
                "message": f"Erro ao obter performance: {str(e)}"
            }
        )


# Incluir rotas
from routes.webhook import webhook_router
app.include_router(webhook_router)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, reload=DEBUG)
