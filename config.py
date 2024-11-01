import datetime

# Grades padrão
grades = {
    "Grade1": {
        "quantidades": {"P": 1, "M": 2, "G": 3, "GG": 2},
        "aproveitamento": 0.88,
        "custo_setup": 250,
        "larguras": {"P": 0.2, "M": 0.22, "G": 0.23, "GG": 0.24},
        "areas": {
            "P": 0.2 * 0.7,
            "M": 0.22 * 0.8,
            "G": 0.23 * 0.9,
            "GG": 0.24 * 1.0,
        },
        "perimetros": {"P": 2.0, "M": 2.2, "G": 2.4, "GG": 2.8},
        "largura_tecido": 1.5,
        "custo_tecido": 18.90,
        "custo_corte": 0.45,
        "custo_enfesto_fixo": 1.25,
        "custo_enfesto_variavel": 0.30,
        "tempo_enfesto_por_metro": 0.5,  # horas por metro de enfesto
        "tempo_corte_por_metro": 0.01,  # horas por metro de perímetro cortado
    },
    "Grade2": {
        "quantidades": {"P": 2, "M": 1, "G": 2, "GG": 1},
        "aproveitamento": 0.87,
        "custo_setup": 300,
        "larguras": {"P": 0.2, "M": 0.22, "G": 0.23, "GG": 0.24},
        "areas": {
            "P": 0.2 * 0.7,
            "M": 0.22 * 0.8,
            "G": 0.23 * 0.9,
            "GG": 0.24 * 1.0,
        },
        "perimetros": {"P": 2.0, "M": 2.2, "G": 2.4, "GG": 2.8},
        "largura_tecido": 1.5,
        "custo_tecido": 18.90,
        "custo_corte": 0.45,
        "custo_enfesto_fixo": 1.25,
        "custo_enfesto_variavel": 0.30,
        "tempo_enfesto_por_metro": 0.5,  # horas por metro de enfesto
        "tempo_corte_por_metro": 0.01,  # horas por metro de perímetro cortado
    },
    "Grade3": {
        "quantidades": {"P": 0, "M": 1, "G": 1, "GG": 0},
        "aproveitamento": 0.85,
        "custo_setup": 200,
        "larguras": {"P": 0.2, "M": 0.22, "G": 0.23, "GG": 0.24},
        "areas": {
            "P": 0.2 * 0.7,
            "M": 0.22 * 0.8,
            "G": 0.23 * 0.9,
            "GG": 0.24 * 1.0,
        },
        "perimetros": {"P": 2.0, "M": 2.2, "G": 2.4, "GG": 2.8},
        "largura_tecido": 1.5,
        "custo_tecido": 18.90,
        "custo_corte": 0.45,
        "custo_enfesto_fixo": 1.25,
        "custo_enfesto_variavel": 0.30,
        "tempo_enfesto_por_metro": 0.5,  # horas por metro de enfesto
        "tempo_corte_por_metro": 0.01,  # horas por metro de perímetro cortado
    },
    "Grade4": {
        "quantidades": {"P": 2, "M": 0, "G": 0, "GG": 1},
        "aproveitamento": 0.86,
        "custo_setup": 180,
        "larguras": {"P": 0.2, "M": 0.22, "G": 0.23, "GG": 0.24},
        "areas": {
            "P": 0.2 * 0.7,
            "M": 0.22 * 0.8,
            "G": 0.23 * 0.9,
            "GG": 0.24 * 1.0,
        },
        "perimetros": {"P": 2.0, "M": 2.2, "G": 2.4, "GG": 2.8},
        "largura_tecido": 1.5,
        "custo_tecido": 18.90,
        "custo_corte": 0.45,
        "custo_enfesto_fixo": 1.25,
        "custo_enfesto_variavel": 0.30,
        "tempo_enfesto_por_metro": 0.5,  # horas por metro de enfesto
        "tempo_corte_por_metro": 0.01,  # horas por metro de perímetro cortado
    },
    "Grade5": {
        "quantidades": {"P": 1, "M": 2, "G": 1, "GG": 2},
        "aproveitamento": 0.86,
        "custo_setup": 280,
        "larguras": {"P": 0.2, "M": 0.22, "G": 0.23, "GG": 0.24},
        "areas": {
            "P": 0.2 * 0.7,
            "M": 0.22 * 0.8,
            "G": 0.23 * 0.9,
            "GG": 0.24 * 1.0,
        },
        "perimetros": {"P": 2.0, "M": 2.2, "G": 2.4, "GG": 2.8},
        "largura_tecido": 1.5,
        "custo_tecido": 18.90,
        "custo_corte": 0.45,
        "custo_enfesto_fixo": 1.25,
        "custo_enfesto_variavel": 0.30,
        "tempo_enfesto_por_metro": 0.5,  # horas por metro de enfesto
        "tempo_corte_por_metro": 0.01,  # horas por metro de perímetro cortado
    },
    "Grade6": {
        "quantidades": {"P": 1, "M": 1, "G": 1, "GG": 1},
        "aproveitamento": 0.87,
        "custo_setup": 220,
        "larguras": {"P": 0.2, "M": 0.22, "G": 0.23, "GG": 0.24},
        "areas": {
            "P": 0.2 * 0.7,
            "M": 0.22 * 0.8,
            "G": 0.23 * 0.9,
            "GG": 0.24 * 1.0,
        },
        "perimetros": {"P": 2.0, "M": 2.2, "G": 2.4, "GG": 2.8},
        "largura_tecido": 1.5,
        "custo_tecido": 18.90,
        "custo_corte": 0.45,
        "custo_enfesto_fixo": 1.25,
        "custo_enfesto_variavel": 0.30,
        "tempo_enfesto_por_metro": 0.5,  # horas por metro de enfesto
        "tempo_corte_por_metro": 0.01,  # horas por metro de perímetro cortado
    },
    "Grade7": {
        "quantidades": {"P": 0, "M": 2, "G": 2, "GG": 1},
        "aproveitamento": 0.88,
        "custo_setup": 270,
        "larguras": {"P": 0.2, "M": 0.22, "G": 0.23, "GG": 0.24},
        "areas": {
            "P": 0.2 * 0.7,
            "M": 0.22 * 0.8,
            "G": 0.23 * 0.9,
            "GG": 0.24 * 1.0,
        },
        "perimetros": {"P": 2.0, "M": 2.2, "G": 2.4, "GG": 2.8},
        "largura_tecido": 1.5,
        "custo_tecido": 18.90,
        "custo_corte": 0.45,
        "custo_enfesto_fixo": 1.25,
        "custo_enfesto_variavel": 0.30,
        "tempo_enfesto_por_metro": 0.5,  # horas por metro de enfesto
        "tempo_corte_por_metro": 0.01,  # horas por metro de perímetro cortado
    },
}

# Configurações padrão dos turnos
turnos_padrao = [
    {"inicio": datetime.time(6, 0), "fim": datetime.time(14, 0), "eficiencia": 1.0},
    {"inicio": datetime.time(14, 0), "fim": datetime.time(22, 0), "eficiencia": 0.9},
    {"inicio": datetime.time(22, 0), "fim": datetime.time(6, 0), "eficiencia": 0.8},
]

# Recursos padrão
recursos_padrao = {
    "enfestadeiras": [
        {"id": "E1", "eficiencia": 0.9},
        {"id": "E2", "eficiencia": 0.93},
        {"id": "E3", "eficiencia": 0.5},
    ],
    "maquinas_corte": [
        {"id": "C1", "eficiencia": 0.95},
        {"id": "C2", "eficiencia": 0.92},
    ],
}
