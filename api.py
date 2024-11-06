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

# Input parameters model
class OptimizationConfig(BaseModel):
    general_configuration: dict
    layouts: list
    fabrics: list
    pieces: list

@app.post("/optimize")
async def optimize_production(
    parameters: OptimizationConfig, 
    current_user: User = Depends(get_current_user)
):
    try:
        input_data = {
            "general_configuration": parameters.general_configuration,
            "layouts": parameters.layouts,
            "fabrics": parameters.fabrics,
            "pieces": parameters.pieces  #
        }
        
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

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/")
async def root():
    return {
        "message": "Production Optimization API",
        "version": "1.0",
        "endpoints": [
            "/token - POST - Authentication",
            "/optimize - POST - Run production optimization",
            "/ - GET - API Information"
        ]
    }