# API de Otimização de Produção

API para otimização de produção têxtil usando algoritmos de seleção de grades.

## Instalação

1. Clone o repositório:

bash
git clone https://github.com/seu-usuario/api-opt-select-grid.git
cd api-opt-select-grid

2. Crie um ambiente virtual e instale as dependências:
bash
python -m venv venv
source venv/bin/activate # Linux/Mac
venv\Scripts\activate # Windows
pip install -r requirements.txt

3. Configure as variáveis de ambiente:
- Copie o arquivo `.env.example` para `.env`
- Edite o arquivo `.env` com suas configurações

## Uso

1. Inicie o servidor:
bash
uvicorn api:app --reload


2. Acesse a documentação:
- http://localhost:8000/docs

## Endpoints

- `POST /token` - Autenticação
- `POST /otimizar` - Executa otimização
- `GET /` - Informações da API



## Estrutura do projeto:


api-opt-select-grid/
├── .env
├── .gitignore
├── README.md
├── requirements.txt
├── api.py
├── auth.py
├── select_grids.py
└── dados_entrada.json