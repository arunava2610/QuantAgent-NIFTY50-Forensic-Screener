from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from nifty_screener import router as nifty_router
from supply_chain_swarm.router import router as sc_router

app = FastAPI(title="Enterprise Agentic API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Attach the decoupled modules
app.include_router(nifty_router)
app.include_router(sc_router)