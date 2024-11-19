from datetime import timedelta
from fastapi import FastAPI, HTTPException, Depends, status
from fastapi.security import OAuth2PasswordRequestForm
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware
from typing import Dict, List

from select_grids_layers import LayoutOptimizer
from auth import (
    User,
    Token, 
    authenticate_user,
    create_access_token,
    get_current_user,
    ACCESS_TOKEN_EXPIRE_MINUTES,
    users_db
)

app = FastAPI(title="Production Optimization API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/token", response_model=Token)
async def login_for_access_token(form_data: OAuth2PasswordRequestForm = Depends()):
    user = authenticate_user(users_db, form_data.username, form_data.password)
    if not user:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect username or password",
            headers={"WWW-Authenticate": "Bearer"},
        )
    access_token_expires = timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    access_token = create_access_token(
        data={"sub": user.username}, expires_delta=access_token_expires
    )
    return {"access_token": access_token, "token_type": "bearer"}

from pydantic import BaseModel, Field
from typing import Dict, List, Optional

class GeneralConfiguration(BaseModel):
    optimality_gap: float = Field(default=0.01, ge=0, le=1)
    solver_time_limit: int = Field(default=6, ge=1)
    max_memory_mb: int = Field(default=2048, ge=1)
    criteria: str = Field(default="waste")
    overproduction_percentage: float = Field(default=0.05, ge=0)
    max_layers: int = Field(default=50, ge=1)
    max_total_length: int = Field(default=10000, ge=1)
    waste_cost: float = Field(default=1000, ge=0)
    layout_change_penalty: float = Field(default=1000, ge=0)
    min_layers_per_layout: int = Field(default=5, ge=1)
    waste_penalty_factor: float = Field(default=2.0, ge=0)

class Layout(BaseModel):
    id: int
    utilization: float
    fabric_width: int
    fabric: str
    layout_length: int
    total_perimeter: int
    utilized_area: int
    waste_area: int
    total_area: int
    pieces: List[Dict]

class Fabric(BaseModel):
    fabric: str
    cost_per_cut_meter: float
    price_per_linear_meter: float
    cost_per_layout_meter: float
    cost_per_layer: float
    fabric_width: int
    max_layers: int

class Piece(BaseModel):
    pattern: str
    fabrics: List[str]
    quantity: Dict[str, int]

class OptimizationConfig(BaseModel):
    general_configuration: GeneralConfiguration
    layouts: List[Layout]
    fabrics: List[Fabric]
    pieces: List[Piece]


@app.post("/optimize")
async def optimize_production(
    parameters: OptimizationConfig, 
    current_user: User = Depends(get_current_user)
):
    try:
        input_data = parameters.dict()
        
        optimizer = LayoutOptimizer(input_data)
        results = optimizer.optimize_production()  

        if any(results.values()):
            optimizer.export_results(results)
            optimizer.export_results_json(results)

        return {
            "status": "success",
            "message": "Optimization completed successfully",
            "data": results
        }

    except ValueError as ve:
        raise HTTPException(status_code=422, detail=str(ve))
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/")
async def root():
    return {
        "message": "Production Optimization API",
        "version": "1.0",
        "parameters": {
            "layout_change_penalty": "Penalidade por usar layouts diferentes (default: 1000)",
            "min_layers_per_layout": "Número mínimo de camadas por layout (default: 5)",
            "waste_penalty_factor": "Fator de multiplicação do custo de desperdício (default: 2.0)"
        },
        "endpoints": [
            "/token - POST - Authentication",
            "/optimize - POST - Run production optimization",
            "/ - GET - API Information"
        ]
    }