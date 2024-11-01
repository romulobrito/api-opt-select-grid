from datetime import timedelta
import datetime
import random

from fastapi import FastAPI, HTTPException, Depends, status
from fastapi.security import OAuth2PasswordRequestForm
from pydantic import BaseModel
from fastapi.middleware.cors import CORSMiddleware
from typing import Dict, List

from min_cost_production import (
    main,
    Turno,
    gerar_pedidos_para_intervalo
)
from auth import (
    User,
    Token, 
    authenticate_user,
    create_access_token,
    get_current_user,
    ACCESS_TOKEN_EXPIRE_MINUTES,
    users_db
)

from config import grades, turnos_padrao, recursos_padrao

app = FastAPI(title="API de Otimização de Produção")

class RecursoModel(BaseModel):
    id: str
    eficiencia: float

class ConfiguracaoOtimizacao(BaseModel):
    criterio: str
    num_dias: int
    tolerancia_largura: float
    percentual_superproducao: float
    max_camadas_por_grade: int
    horas_producao: float
    comprimento_mesa_enfesto: float
    penalizacao_superproducao: float
    relaxacao: bool

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

from fastapi.security import OAuth2PasswordRequestForm

@app.post("/token", response_model=Token)
async def login_for_access_token(form_data: OAuth2PasswordRequestForm = Depends()):
    user = authenticate_user(users_db, form_data.username, form_data.password)
    if not user:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Usuário ou senha incorretos",
            headers={"WWW-Authenticate": "Bearer"},
        )
    access_token_expires = timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    access_token = create_access_token(
        data={"sub": user.username}, expires_delta=access_token_expires
    )
    return {"access_token": access_token, "token_type": "bearer"}


@app.post("/otimizar")
async def otimizar_producao(parametros_otimizacao: ConfiguracaoOtimizacao, current_user: User = Depends(get_current_user)):
    try:
        # Configurações padrão
        data_inicio = datetime.datetime.now()
        
        # Converter turnos do config para objetos Turno
        turnos = [
            Turno(t["inicio"], t["fim"], eficiencia=t["eficiencia"])
            for t in turnos_padrao
        ]

        # Gerar pedidos de teste
        pedidos = gerar_pedidos_para_intervalo(data_inicio, parametros_otimizacao.num_dias)

        # Executar otimização
        resultado = main(
            criterio_prioridade=parametros_otimizacao.criterio,
            data_inicio=data_inicio,
            num_dias=parametros_otimizacao.num_dias,
            recursos=recursos_padrao,
            tolerancia_largura=parametros_otimizacao.tolerancia_largura,
            percentual_superproducao=parametros_otimizacao.percentual_superproducao,
            max_camadas_por_grade=parametros_otimizacao.max_camadas_por_grade,
            grades=grades,
            pedidos=pedidos,
            turnos=turnos,
            horas_producao=parametros_otimizacao.horas_producao,
            comprimento_mesa_enfesto=parametros_otimizacao.comprimento_mesa_enfesto,
            penalizacao_superproducao=parametros_otimizacao.penalizacao_superproducao,
            relaxacao=parametros_otimizacao.relaxacao,
        )

        return {
            "status": "success",
            "message": "Otimização concluída com sucesso",
            "data": resultado
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# Função auxiliar para gerar pedidos
# def gerar_pedidos_para_intervalo(data_inicio, num_dias):
#     pedidos = {}
#     tamanhos = ["P", "M", "G", "GG"]
    
#     for i in range(1, 4):  # Gerar 3 pedidos de teste
#         prazo = data_inicio + datetime.timedelta(days=random.randint(1, num_dias))
        
#         # Gerar demandas aleatórias para cada tamanho
#         demandas = {
#             tamanho: random.randint(50, 200) 
#             for tamanho in tamanhos
#         }
        
#         pedidos[f"P{i}"] = {
#             "prazo": prazo,
#             "demandas": demandas,
#             "prioridade": random.randint(1, 7)
#         }
    
#     if not pedidos:
#         raise ValueError("Nenhum pedido foi gerado")
        
#     return pedidos
def gerar_demanda_flutuante(demanda_base, variacao=0.05):
    random.seed(10)  # Reproducibilidade
    return int(demanda_base * (1 + random.uniform(-variacao, variacao)))
def gerar_pedidos_para_intervalo(data_inicio, num_dias, pedidos_reais=None):
    if pedidos_reais:
        return pedidos_reais

    random.seed(10)  # Reproducibilidade
    pedidos = {}
    pedido_id = 1
    for dia in range(num_dias):
        data_atual = data_inicio + datetime.timedelta(days=dia)
        days1 = random.randint(1, 7)
        days2 = random.randint(1, 7)
        days3 = random.randint(1, 7)
        novos_pedidos = {
            pedido_id: {
                "demandas": {
                    t: gerar_demanda_flutuante(d, variacao=0.05)
                    for t, d in {"P": 50, "M": 100, "G": 80, "GG": 70}.items()
                },
                "prazo": data_atual + datetime.timedelta(days=days1),
            },
            pedido_id
            + 1: {
                "demandas": {
                    t: gerar_demanda_flutuante(d, variacao=0.05)
                    for t, d in {"P": 30, "M": 60, "G": 90, "GG": 40}.items()
                },
                "prazo": data_atual + datetime.timedelta(days=days2),
            },
            pedido_id
            + 2: {
                "demandas": {
                    t: gerar_demanda_flutuante(d, variacao=0.05)
                    for t, d in {"P": 40, "M": 100, "G": 100, "GG": 80}.items()
                },
                "prazo": data_atual + datetime.timedelta(days=days3),
            },
        }
        pedidos.update(novos_pedidos)
        pedido_id += 3
    return pedidos

@app.get("/")
async def root():
    return {
        "message": "API de Otimização de Produção",
        "version": "1.0",
        "endpoints": [
            "/otimizar - POST - Executa otimização de produção",
            "/ - GET - Informações da API"
        ]
    }

