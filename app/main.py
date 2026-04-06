"""
Sistema de Trading Semi-Automático - API Principal
"""

from fastapi import FastAPI
from fastapi.responses import JSONResponse
from config import DEBUG

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


# Incluir rotas
from routes.webhook import webhook_router
app.include_router(webhook_router)


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, reload=DEBUG)
