import requests
import json
from datetime import datetime
from tabulate import tabulate



def formatar_data(data_str):
    return datetime.fromisoformat(data_str).strftime("%d/%m/%Y %H:%M")

def formatar_moeda(valor):
    return f"R$ {valor:,.2f}"

def get_token():
    """Function to obtain authentication token"""
    login_url = "http://localhost:8000/token"
    
    data = {
        "username": "admin",
        "password": "admin123",
        "grant_type": "password"  
    }
    
    try:
        response = requests.post(
            login_url,
            data=data,  
            headers={"Content-Type": "application/x-www-form-urlencoded"}
        )
        
        if response.status_code == 200:
            return response.json()["access_token"]
        else:
            print(f"Error getting token: {response.status_code}")
            print(f"Response: {response.text}")
            return None
            
    except Exception as e:
        print(f"Error getting token: {e}")
        return None

def view_results():
    print("Starting visualization script...")
    token = get_token()
    if not token:
        print("Could not obtain authentication token")
        return

    url = "http://localhost:8000/optimize"
    
    try:
        with open('dados_entrada.json', 'r', encoding='utf-8') as f:  
            data = json.load(f)
        
        # Validação básica dos dados
        required_keys = ['general_configuration', 'layouts', 'fabrics', 'pieces']
        for key in required_keys:
            if key not in data:
                raise ValueError(f"Missing required key: {key}")
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }

        print(f"Sending request to {url}...")
        response = requests.post(url, json=data, headers=headers)
        print(f"Response status: {response.status_code}")

        if response.status_code == 422:
            error_detail = response.json().get('detail', 'Unknown validation error')
            print(f"Validation error: {error_detail}")
            return
        elif response.status_code != 200:
            print(f"Request error: {response.text}")
            return

        results = response.json()
        
        if results["status"] == "success":
            print("\nOptimization Results:")
            for order_id, result in results["data"].items():
                if result:
                    print(f"\nPeça: {result['pattern']}")
                    print("Produção por tamanho:")
                    for size, qty in result['production'].items():
                        print(f"  {size}: {qty}")
                    print("Métricas:")
                    metrics = result['metrics']
                    print(f"  Custo total: {formatar_moeda(metrics['total_cost'])}")
                    print(f"  Tecido total utilizado: {metrics['fabric_meters']:.2f} metros lineares")
                    print(f"  Desperdício: {metrics['fabric_waste_area']:.2f} m²")
                    print(f"  Desperdício: {metrics['fabric_waste_meters']:.2f} metros lineares")
                else:
                    print(f"\nPeça {order_id}: Sem solução viável")

    except requests.exceptions.ConnectionError:
        print("\nErro: Não foi possível conectar à API.")
    except Exception as e:
        print(f"\nErro inesperado: {e}")
        import traceback
        print(traceback.format_exc())

if __name__ == "__main__":
    view_results()