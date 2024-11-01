import requests
import json
from datetime import datetime
from tabulate import tabulate

def formatar_data(data_str):
    return datetime.fromisoformat(data_str).strftime("%d/%m/%Y %H:%M")

def formatar_moeda(valor):
    return f"R$ {valor:,.2f}"

# def get_token():
#     """Função para obter o token de autenticação"""
#     login_url = "http://localhost:8000/token"
#     credentials = {
#         "username": "admin",
#         "password": "admin123"
#     }
#     response = requests.post(
#         login_url,
#         data=credentials
#     )
#     return response.json()["access_token"]


def get_token():
    """Função para obter o token de autenticação"""
    login_url = "http://localhost:8000/token"
    
    data = {
        "username": "admin",
        "password": "admin123"
    }
    
    try:
        response = requests.post(
            login_url,
            data=data  # Enviar como form-data
        )
        
        if response.status_code == 200:
            return response.json()["access_token"]
        else:
            print(f"Erro ao obter token: {response.status_code}")
            print(f"Resposta: {response.text}")
            return None
            
    except Exception as e:
        print(f"Erro ao obter token: {e}")
        return None


def visualizar_resultados():
    print("Iniciando script de visualização...")
    token = get_token()
    if not token:
        print("Não foi possível obter o token de autenticação")
        return
    url = "http://localhost:8000/otimizar"
    data = {
        "criterio": "prazo",
        "num_dias": 3,
        "tolerancia_largura": 0.05,
        "percentual_superproducao": 0.05,
        "max_camadas_por_grade": 30,
        "horas_producao": 24,
        "comprimento_mesa_enfesto": 10.0,
        "penalizacao_superproducao": 10000,
        "relaxacao": False
    }
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    try:
        print(f"Enviando requisição para {url}...")
        response = requests.post(url, json=data, headers=headers)
        print(f"Status da resposta: {response.status_code}")
        
        if response.status_code != 200:
            print(f"Erro na requisição: {response.text}")
            return

        resultados = response.json()
        print("Dados recebidos com sucesso!")

        # Cabeçalho
        print("\n" + "="*50)
        print("RELATÓRIO DE OTIMIZAÇÃO DE PRODUÇÃO")
        print("="*50)
        
        # Parâmetros da Otimização
        print("\nParâmetros Utilizados:")
        params = [
            ["Critério", data["criterio"]],
            ["Número de Dias", data["num_dias"]],
            ["Tolerância Largura", f"{data['tolerancia_largura']*100}%"],
            ["Superprodução", f"{data['percentual_superproducao']*100}%"],
            ["Máx. Camadas/Grade", data["max_camadas_por_grade"]],
            ["Horas Produção", data["horas_producao"]],
            ["Comp. Mesa Enfesto", f"{data['comprimento_mesa_enfesto']}m"],
        ]
        print(tabulate(params, tablefmt="grid"))

        # Métricas Globais
        print("\nMétricas Globais:")
        metricas = resultados["data"]["metricas_globais"]
        metricas_table = [
            ["Custo Total", formatar_moeda(metricas["custo_total"])],
            ["Tempo Total", f"{metricas['tempo_total']:.2f} horas"],
            ["Desperdício Total", f"{metricas['desperdicio_total']:.2f}%"]
        ]
        print(tabulate(metricas_table, tablefmt="grid"))

        # Cronograma de Produção
        print("\nCronograma de Produção:")
        cronograma_data = []
        for pedido, info in resultados["data"]["cronograma"].items():
            cronograma_data.append([
                f"Pedido {pedido}",
                "Enfestamento",
                formatar_data(info["inicio_enfestamento"]),
                formatar_data(info["fim_enfestamento"]),
                info["enfestadeira"]
            ])
            cronograma_data.append([
                f"Pedido {pedido}",
                "Corte",
                formatar_data(info["inicio_corte"]),
                formatar_data(info["fim_corte"]),
                info["maquina_corte"]
            ])
        print(tabulate(cronograma_data, 
                      headers=["Pedido", "Operação", "Início", "Fim", "Recurso"],
                      tablefmt="grid"))

        # Detalhes dos Pedidos
        print("\nDetalhes dos Pedidos:")
        for pedido, info in resultados["data"]["resultados"].items():
            print(f"\nPedido {pedido}:")
            detalhes = [
                ["Prazo", formatar_data(info["prazo"])],
                ["Custo Total", formatar_moeda(info["custo_total"])],
                ["Custo Setup", formatar_moeda(info["custo_setup"])],
                ["Metros de Tecido", f"{info['metros_tecido']:.2f}m"],
                ["Perímetro Cortado", f"{info['perimetro_cortado']:.2f}m"],
                ["Desperdício", f"{info['desperdicio']:.2f}%"]
            ]
            print(tabulate(detalhes, tablefmt="simple"))
            
            print("\nProdução x Demanda:")
            prod_data = []
            for tamanho in info["producao"].keys():
                prod_data.append([
                    tamanho,
                    info["producao"][tamanho],
                    info["demandas"][tamanho],
                    info["producao"][tamanho] - info["demandas"][tamanho]
                ])
            print(tabulate(prod_data, 
                         headers=["Tamanho", "Produção", "Demanda", "Diferença"],
                         tablefmt="simple"))

        # Salvar resultados em arquivo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        arquivo_saida = f'relatorio_otimizacao_{timestamp}.txt'
        
        print(f"\nSalvando resultados em {arquivo_saida}...")
        with open(arquivo_saida, 'w', encoding='utf-8') as f:
            f.write("RELATÓRIO DE OTIMIZAÇÃO DE PRODUÇÃO\n")
            f.write("="*50 + "\n\n")
            
            # Salvar métricas globais
            f.write("Métricas Globais:\n")
            f.write(f"Custo Total: {formatar_moeda(metricas['custo_total'])}\n")
            f.write(f"Tempo Total: {metricas['tempo_total']:.2f} horas\n")
            f.write(f"Desperdício Total: {metricas['desperdicio_total']:.2f}%\n\n")
            
            # Salvar cronograma
            f.write("Cronograma de Produção:\n")
            f.write(tabulate(cronograma_data, 
                           headers=["Pedido", "Operação", "Início", "Fim", "Recurso"],
                           tablefmt="grid"))
            f.write("\n\n")
            
            # Salvar detalhes dos pedidos
            f.write("Detalhes dos Pedidos:\n")
            for pedido, info in resultados["data"]["resultados"].items():
                f.write(f"\nPedido {pedido}:\n")
                f.write(tabulate(detalhes, tablefmt="simple"))
                f.write("\n\nProdução x Demanda:\n")
                f.write(tabulate(prod_data, 
                               headers=["Tamanho", "Produção", "Demanda", "Diferença"],
                               tablefmt="simple"))
                f.write("\n")
        
        print(f"Relatório salvo com sucesso em {arquivo_saida}")

    except requests.exceptions.ConnectionError:
        print("\nErro: Não foi possível conectar à API. Verifique se o servidor está rodando.")
    except Exception as e:
        print(f"\nErro inesperado: {e}")
        import traceback
        print(traceback.format_exc())

if __name__ == "__main__":
    visualizar_resultados()