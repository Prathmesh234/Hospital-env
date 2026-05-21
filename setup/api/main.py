"""FastAPI application factory."""

from __future__ import annotations

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from setup import __version__
from setup.api.routers import (
    billing,
    clinical,
    encounters,
    health,
    medications,
    patients,
    providers,
)


def create_app() -> FastAPI:
    app = FastAPI(
        title="Hospital-env API",
        description="A complex synthetic hospital information system, backed by PostgreSQL.",
        version=__version__,
    )

    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_methods=["*"],
        allow_headers=["*"],
    )

    app.include_router(health.router)
    app.include_router(patients.router)
    app.include_router(providers.router)
    app.include_router(encounters.router)
    app.include_router(clinical.router)
    app.include_router(medications.router)
    app.include_router(billing.router)

    return app


app = create_app()
