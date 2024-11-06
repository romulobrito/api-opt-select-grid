import json
import logging
from ortools.linear_solver import pywraplp
from openpyxl import Workbook
import unicodedata

# Logging Configuration
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("debug.log"), logging.StreamHandler()],
)

class LayoutOptimizer:
    def __init__(self, input_data):
        self.config = input_data['general_configuration']
        self.layouts = input_data.get('layout', input_data.get('layouts', []))
        self.fabrics = {f['fabric']: f for f in input_data['fabrics']}
        self.pieces = input_data['pieces']  
        self.sizes = ["P", "M", "G", "GG"]
        self.overproduction_penalty = self.config.get('overproduction_percentage', 0.05)
        self.unit_waste_cost = self.config.get('waste_cost', 0.1)
        self.optimality_gap = self.config.get('optimality_gap', 0.01) 
        self.max_memory_mb = self.config.get('max_memory_mb', 2048)



    def optimize_production(self):
        """Método principal de otimização"""
        results = {}
        
        for i, piece in enumerate(self.pieces, 1):
            order_id = f'order_{i}'
            logging.info(f"Processando {piece['pattern']} (Ordem {order_id})")
            
            # Cria estrutura de ordem sintética
            order = {
                'id': order_id,
                'pieces': [piece],
                'fabric_width': self.fabrics[piece['fabrics'][0]]['fabric_width'],
                'max_layers': self.fabrics[piece['fabrics'][0]]['max_layers'],
                'max_length': self.config['max_total_length']
            }
            
            result = self.optimize_order(order)
            if result:
                results[order_id] = result
                logging.info(f"Solução encontrada para {piece['pattern']}")
            else:
                logging.warning(f"Não foi possível encontrar solução para {piece['pattern']}")
                
        return results


    def calculate_layout_costs(self, layout, num_layers, fabric_cost):
        """Calcula custos do layout"""
        fabric = self.fabrics[layout['fabric']]
        
        return {
            'cutting_cost': fabric['cost_per_cut_meter'] * layout['total_perimeter'] * num_layers / 1000,
            'layout_cost': (
                fabric['cost_per_layer'] * num_layers +
                fabric['cost_per_layout_meter'] * layout['layout_length'] * num_layers / 1000
            ),
            'fabric_cost': fabric['price_per_linear_meter'] * layout['layout_length'] * num_layers / 1000,
            'waste_cost': self.unit_waste_cost * layout['waste_area'] * num_layers / 1_000_000
        }

    def preprocess_layouts(self, demand, fabric_width):
        """Preprocess and filter layouts compatible with the demand and fabric width"""
        filtered_layouts = []
        for layout in self.layouts:
            # Filter by fabric width
            if layout['fabric_width'] != fabric_width:
                continue
            
            # Check if the layout contains the necessary patterns
            layout_patterns = [p['pattern'] for p in layout['pieces']]
            order_patterns = [p['pattern'] for p in demand['pieces']]
            if not set(order_patterns).issubset(set(layout_patterns)):
                continue
            
            # Calculate efficiency per size
            layout['efficiency'] = {}
            for size in self.sizes:
                total_pieces_size = sum(
                    p['size_grade'].get(size, 0) for p in layout['pieces']
                )
                pieces_per_meter = total_pieces_size / (layout['layout_length'] / 1000)
                efficiency = pieces_per_meter * layout['utilization']
                layout['efficiency'][size] = efficiency

            filtered_layouts.append(layout)
                
        return filtered_layouts
    
    # def _process_solution(self, solver, layouts, x, demand, pattern, piece, order):
    #     """Processa a solução do solver"""
    #     try:
    #         result = {
    #             'status': 'optimal',
    #             'pattern': pattern,
    #             'production': {},
    #             'overproduction': {},
    #             'layouts_used': [],
    #             'pieces': [piece],  # Adiciona a peça aqui
    #             'metrics': {
    #                 'total_cost': solver.Objective().Value(),
    #                 'fabric_meters': 0,
    #                 'fabric_waste': 0,
    #                 'cutting_cost': 0,
    #                 'layout_cost': 0
    #             }
    #         }

    #         # Processa layouts utilizados
    #         for layout in layouts:
    #             num_layers = int(x[layout['id']].solution_value())
    #             if num_layers > 0:
    #                 # Calcula produção por tamanho para este layout
    #                 production_per_size = {}
    #                 for size in self.sizes:
    #                     if size in layout['pieces'][0]['size_grade']:
    #                         qty_per_layer = layout['pieces'][0]['size_grade'][size]
    #                         production_per_size[size] = qty_per_layer * num_layers
                    
    #                 logging.info(f"\nLayout {layout['id']} selecionado:")
    #                 logging.info(f"  Número de camadas: {num_layers}")
    #                 logging.info(f"  Comprimento: {layout['layout_length']/1000:.2f}m")
    #                 logging.info(f"  Aproveitamento: {layout['utilization']*100:.1f}%")
    #                 logging.info("  Peças por camada:")
    #                 for size, qty in layout['pieces'][0]['size_grade'].items():
    #                     logging.info(f"    {size}: {qty}")
    #                 logging.info(f"  Total de peças produzidas:")
    #                 for size, qty in production_per_size.items():
    #                     logging.info(f"    {size}: {qty}")

    #                 # Calcula custos
    #                 fabric_cost = (layout['layout_length'] * num_layers * 
    #                             self.fabrics[piece['fabrics'][0]]['price_per_linear_meter'] / 1000)
    #                 cutting_cost = (layout['total_perimeter'] * num_layers * 
    #                             self.fabrics[piece['fabrics'][0]]['cost_per_cut_meter'] / 1000)
    #                 layout_cost = num_layers * self.fabrics[piece['fabrics'][0]]['cost_per_layer']

    #                 result['layouts_used'].append({
    #                     'layout_id': layout['id'],
    #                     'layers': num_layers,
    #                     'length': layout['layout_length'],
    #                     'utilization': layout['utilization'],
    #                     'production_per_size': production_per_size,
    #                     'costs': {
    #                         'fabric': fabric_cost,
    #                         'cutting': cutting_cost,
    #                         'layout': layout_cost
    #                     }
    #                 })

    #                 # Atualiza métricas
    #                 result['metrics']['fabric_meters'] += (layout['layout_length'] * num_layers / 1000)
    #                 result['metrics']['fabric_waste'] += (layout['waste_area'] * num_layers)
    #                 result['metrics']['cutting_cost'] += cutting_cost
    #                 result['metrics']['layout_cost'] += layout_cost

    #         # Calcula produção total por tamanho
    #         for size in self.sizes:
    #             result['production'][size] = sum(
    #                 layout['production_per_size'].get(size, 0)
    #                 for layout in result['layouts_used']
    #             )
                
    #             if size in demand:
    #                 diff = result['production'][size] - demand[size]
    #                 logging.info(f"\nTamanho {size}:")
    #                 logging.info(f"  Demanda: {demand[size]}")
    #                 logging.info(f"  Produção: {result['production'][size]}")
    #                 logging.info(f"  Diferença: {diff:+d}")

    #         return result
            
    #     except Exception as e:
    #         logging.error(f"Erro ao processar solução: {str(e)}")
    #         import traceback
    #         logging.error(traceback.format_exc())
    #         raise


    def _process_solution(self, solver, layouts, x, demand, pattern, piece, order):
        """Processa a solução do solver"""
        try:
            result = {
                'status': 'optimal',
                'pattern': pattern,
                'demand': piece['quantity'], 
                'production': {},
                'overproduction': {},
                'layouts_used': [],
                'pieces': [piece],
                'metrics': {
                    'fabric_meters': 0,
                    'fabric_cost': 0,
                    'cutting_cost': 0,
                    'layout_setup_cost': 0,
                    'layer_cost': 0,
                    'waste_cost': 0,
                    'total_cost': 0,
                    'fabric_waste': 0
                }
            }

            # Inicializa produção total por tamanho
            total_production = {size: 0 for size in self.sizes}

            # Processa layouts utilizados
            for layout in layouts:
                num_layers = int(x[layout['id']].solution_value())
                if num_layers > 0:
                    fabric = self.fabrics[piece['fabrics'][0]]
                    
                    # Calcula produção por tamanho para este layout
                    production_per_size = {}
                    for size in self.sizes:
                        if size in layout['pieces'][0]['size_grade']:
                            qty_per_layer = layout['pieces'][0]['size_grade'][size]
                            production_per_size[size] = qty_per_layer * num_layers
                            total_production[size] += production_per_size[size]

                    # Calcula métricas do layout
                    layout_meters = layout['layout_length'] * num_layers / 1000
                    waste_area = layout['waste_area'] * num_layers / 1_000_000

                    # Calcula custos
                    layout_costs = self.calculate_layout_costs(layout, num_layers, fabric['price_per_linear_meter'])

                    # Adiciona layout usado
                    layout_info = {
                        'layout_id': layout['id'],
                        'layers': num_layers,
                        'length': layout['layout_length'],
                        'utilization': layout['utilization'],
                        'waste_area': layout['waste_area'],
                        'production_per_size': production_per_size,
                        'costs': layout_costs
                    }
                    result['layouts_used'].append(layout_info)

                    # Atualiza métricas totais
                    result['metrics']['fabric_meters'] += layout_meters
                    result['metrics']['fabric_cost'] += layout_costs['fabric_cost']
                    result['metrics']['cutting_cost'] += layout_costs['cutting_cost']
                    result['metrics']['layout_setup_cost'] += layout_costs['layout_cost']
                    result['metrics']['layer_cost'] += fabric['cost_per_layer'] * num_layers
                    result['metrics']['waste_cost'] += layout_costs['waste_cost']
                    result['metrics']['fabric_waste'] += waste_area

            # Calcula produção e superprodução
            result['production'] = total_production
            result['overproduction'] = {
                size: max(0, total_production[size] - piece['quantity'][size])
                for size in self.sizes
            }

            # Calcula custo total
            result['metrics']['total_cost'] = (
                result['metrics']['fabric_cost'] +
                result['metrics']['cutting_cost'] +
                result['metrics']['layout_setup_cost'] +
                result['metrics']['layer_cost'] +
                result['metrics']['waste_cost']
            )

            return result

        except Exception as e:
            logging.error(f"Erro ao processar solução: {str(e)}")
            raise

    def optimize_order(self, order):
        """Otimiza uma ordem específica permitindo múltiplos layouts"""
        try:
            piece = order['pieces'][0]
            pattern = piece['pattern']
            demand_quantity = piece['quantity']
            fabric = piece['fabrics'][0]
            
            logging.info(f"Iniciando otimização para {pattern}")
            logging.info(f"Demanda: {demand_quantity}")
            
            # Filtra e ordena layouts
            filtered_layouts = [
                layout for layout in self.layouts
                if layout['fabric'] == fabric 
                and any(p['pattern'] == pattern for p in layout['pieces'])
            ]
            filtered_layouts.sort(key=lambda x: x['utilization'], reverse=True)
            
            # Log dos layouts disponíveis
            for layout in filtered_layouts:
                logging.info(f"Layout {layout['id']}: {layout['utilization']*100:.1f}% utilização")
                logging.info("  Peças por camada:")
                for size, qty in layout['pieces'][0]['size_grade'].items():
                    logging.info(f"    {size}: {qty}")
            
            # Configuração do solver
            solver = pywraplp.Solver.CreateSolver('SCIP')
            solver.SetTimeLimit(self.config.get('solver_time_limit', 1) * 60 * 1000)  # converte minutos para milissegundos
            scip_params = (
                f"limits/gap = {self.optimality_gap}\n"
                f"limits/memory = {self.max_memory_mb}\n"
                "display/verblevel = 4\n"  # Nível de log detalhado
                "timing/clocktype = 1\n"    # Usa tempo de CPU
            )
            solver.SetSolverSpecificParametersAsString(scip_params)
            
            # Variáveis de decisão
            x = {}  # número de camadas por layout
            y = {}  # variável binária indicando se o layout é usado
            
            # Inicializa variáveis para cada layout
            for layout in filtered_layouts:
                x[layout['id']] = solver.IntVar(0, self.config['max_layers'], f'x_{layout["id"]}')
                y[layout['id']] = solver.IntVar(0, 1, f'y_{layout["id"]}')
                
                # Relaciona x e y: se y=0, x deve ser 0
                solver.Add(x[layout['id']] <= self.config['max_layers'] * y[layout['id']])
            
            # Garante uso de pelo menos um layout
            solver.Add(solver.Sum(y[l['id']] for l in filtered_layouts) >= 1)
            
            # Restrições de demanda por tamanho
            for size in self.sizes:
                if size in demand_quantity:
                    # Soma a produção de todos os layouts
                    total_production = solver.Sum([
                        x[layout['id']] * layout['pieces'][0]['size_grade'].get(size, 0)
                        for layout in filtered_layouts
                    ])
                    
                    # Garante produção mínima
                    solver.Add(total_production >= demand_quantity[size])
                    
                    # Limita excesso de produção
                    solver.Add(total_production <= demand_quantity[size] * 1.05)  # 5% máximo
                    
                    logging.info(f"Configurada restrição para tamanho {size}:")
                    logging.info(f"  Demanda: {demand_quantity[size]}")
                    logging.info(f"  Produção mínima: {demand_quantity[size]}")
                    logging.info(f"  Produção máxima: {demand_quantity[size] * 1.05}")
            
            # Função objetivo: minimizar custos totais
            objective = solver.Sum([
                x[layout['id']] * (
                    # Custo do tecido
                    # layout['layout_length'] * self.fabrics[fabric]['price_per_linear_meter'] / 1000 +
                    # Custo de corte
                    layout['total_perimeter'] * self.fabrics[fabric]['cost_per_cut_meter'] / 1000 +
                    # Custo por camada
                    self.fabrics[fabric]['cost_per_layer'] +
                    # Penalidade por desperdício
                    layout['waste_area'] * self.unit_waste_cost / 1_000_000
                )
                for layout in filtered_layouts
            ])
            
            solver.Minimize(objective)
            
            # Resolve o problema
            status = solver.Solve()
            logging.info(f"Status da solução: {status}")
            
            if status == solver.OPTIMAL:
                logging.info("Solução ótima encontrada!")
                logging.info(f"Valor objetivo: {solver.Objective().Value():.6f}")
                logging.info(f"Melhor limite: {solver.Objective().BestBound():.6f}")
                gap = abs(solver.Objective().Value() - solver.Objective().BestBound()) / abs(solver.Objective().Value())
                logging.info(f"Gap de otimalidade: {gap*100:.6f}%")
            elif status == solver.FEASIBLE:
                logging.info("Solução viável encontrada (não necessariamente ótima)")
                logging.info(f"Gap de otimalidade: {gap*100:.6f}%")
            else:
                logging.warning("Nenhuma solução viável encontrada")
                return None
                
            # Log da solução encontrada
            if status in [solver.OPTIMAL, solver.FEASIBLE]:
                logging.info("\nDetalhes da solução:")
                total_production = {size: 0 for size in self.sizes}
                
                for layout in filtered_layouts:
                    num_layers = int(x[layout['id']].solution_value())
                    if num_layers > 0:
                        logging.info(f"\nLayout {layout['id']} selecionado:")
                        logging.info(f"  Número de camadas: {num_layers}")
                        logging.info(f"  Aproveitamento: {layout['utilization']*100:.1f}%")
                        logging.info("  Produção por tamanho:")
                        
                        for size in self.sizes:
                            if size in layout['pieces'][0]['size_grade']:
                                prod = num_layers * layout['pieces'][0]['size_grade'][size]
                                total_production[size] += prod
                                logging.info(f"    {size}: {prod}")
                
                logging.info("\nProdução total por tamanho:")
                for size in self.sizes:
                    if size in demand_quantity:
                        logging.info(f"  {size}: {total_production[size]} (Demanda: {demand_quantity[size]})")
                
                if status == pywraplp.Solver.OPTIMAL or status == pywraplp.Solver.FEASIBLE:
                    result = self._process_solution(
                        solver=solver,
                        layouts=filtered_layouts,
                        x=x,
                        demand=demand_quantity,
                        pattern=pattern,
                        piece=piece,
                        order=order
                    )
                    
                    # Adiciona informações do solver
                    result['solver_info'] = {
                        'status': 'optimal' if status == pywraplp.Solver.OPTIMAL else 'feasible',
                        'gap': gap,
                        'objective_value': solver.Objective().Value(),
                        'best_bound': solver.Objective().BestBound(),
                        'iterations': solver.iterations(),
                        'wall_time': solver.WallTime()/1000.0
                    }
                    
                    return result
                
            return None
            
        except Exception as e:
            logging.error(f"Erro na otimização: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())
            return None
        

    def export_results(self, results):
        """Exporta resultados para Excel"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Resultados"
        
        # Cabeçalhos
        headers = [
            "Ordem", "Padrão", "Layout", "Camadas", 
            "Comprimento (m)", "Aproveitamento (%)",
            "P", "M", "G", "GG",
            "Custo Total", "Metros Tecido", "Desperdício (m²)"
        ]
        ws.append(headers)
        
        # Dados
        for order_id, result in results.items():
            if result:
                for layout in result['layouts_used']:
                    row = [
                        order_id,
                        result['pattern'],
                        layout['layout_id'],
                        layout['layers'],
                        layout['length']/1000,
                        layout['utilization']*100,
                        layout['production_per_size'].get('P', 0),
                        layout['production_per_size'].get('M', 0),
                        layout['production_per_size'].get('G', 0),
                        layout['production_per_size'].get('GG', 0),
                        result['metrics']['total_cost'],
                        result['metrics']['fabric_meters'],
                        result['metrics']['fabric_waste']/1_000_000
                    ]
                    ws.append(row)
        
        # Salva arquivo
        wb.save("results.xlsx")

   

    @staticmethod
    def remove_accents(text):
        """Remove acentos e caracteres especiais de uma string"""
        try:
            text = str(text)
            normalized = unicodedata.normalize('NFKD', text)
            return u"".join([c for c in normalized if not unicodedata.combining(c)])
        except Exception:
            return text

    # def export_results_json(self, results):
    #     """Exporta resultados para JSON"""
    #     try:
    #         output = []
    #         for order_id, result in results.items():
    #             if result:

    #                 # result['metrics']['fabric_waste'] = result['metrics']['fabric_waste'] / 1_000_000
    #                 # Usa get() para evitar KeyError
    #                 pieces = result.get('pieces', [{
    #                     'pattern': result['pattern'],
    #                     'fabrics': [],
    #                     'quantity': {}
    #                 }])
                    
    #                 output_item = {
    #                     "order_id": order_id,
    #                     "pattern": self.remove_accents(result['pattern']),
    #                     "metrics": result['metrics'],
    #                     "production": [{
    #                         "pattern": self.remove_accents(p['pattern']),
    #                         "fabrics": [self.remove_accents(f) for f in p.get('fabrics', [])],
    #                         "quantity": p.get('quantity', {}),
    #                         "production": result.get('production', {}),
    #                         "layout_ids": [l.get('layout_id', l.get('id', 0)) for l in result.get('layouts_used', [])]
    #                     } for p in pieces],
    #                     "layouts": [{
    #                         "id": layout.get('layout_id', layout.get('id', 0)),
    #                         "utilization": layout.get('utilization', 0),
    #                         "fabric_width": layout.get('fabric_width', 0),
    #                         "fabric": self.remove_accents(layout.get('fabric', '')),
    #                         "layout_length": layout.get('layout_length', 0),
    #                         "total_perimeter": layout.get('total_perimeter', 0),
    #                         "utilized_area": layout.get('utilized_area', 0),
    #                         "waste_area": layout.get('waste_area', 0),
    #                         "total_area": layout.get('total_area', 0),
    #                         "pieces": layout.get('pieces', []),
    #                         "efficiency": layout.get('efficiency', {}),
    #                         "layers": layout.get('layers', 0),
    #                         "order_ids": [order_id]
    #                     } for layout in result.get('layouts_used', [])]
    #                 }
    #                 output.append(output_item)
            
    #         with open('results.json', 'w', encoding='utf-8') as f:
    #             json.dump(output, f, indent=4, ensure_ascii=False)
                
    #     except Exception as e:
    #         logging.error(f"Erro ao exportar resultados JSON: {str(e)}")
    #         import traceback
    #         logging.error(traceback.format_exc())
    #         raise

    def export_results_json(self, results):
        """Exporta resultados para JSON"""
        try:
            output = []
            for order_id, result in results.items():
                if result:
                    output_item = {
                        "order_id": order_id,
                        "pattern": self.remove_accents(result['pattern']),
                        "metrics": {
                            "fabric_meters": result['metrics']['fabric_meters'],
                            "fabric_cost": result['metrics']['fabric_cost'],
                            "cutting_cost": result['metrics']['cutting_cost'],
                            "layout_setup_cost": result['metrics']['layout_setup_cost'],
                            "layer_cost": result['metrics']['layer_cost'],
                            "waste_cost": result['metrics']['waste_cost'],
                            "total_cost": result['metrics']['total_cost'],
                            "fabric_waste": result['metrics']['fabric_waste']
                        },
                        "demand": {
                            size: qty for size, qty in result.get('demand', {}).items()
                        },
                        "production": {
                            size: qty for size, qty in result.get('production', {}).items()
                        },
                        "overproduction": {
                            size: qty for size, qty in result.get('overproduction', {}).items()
                        },
                        "layouts_used": [
                            {
                                "layout_id": layout['layout_id'],
                                "num_layers": layout['layers'],
                                "length_meters": layout['length'] / 1000,
                                "utilization": layout['utilization'],
                                "waste_area": layout['waste_area'] / 1_000_000,  # mm² para m²
                                "production_per_size": layout['production_per_size'],
                                "costs": {
                                    "fabric_cost": layout['costs']['fabric_cost'],
                                    "cutting_cost": layout['costs']['cutting_cost'],
                                    "layout_cost": layout['costs']['layout_cost']
                                }
                            }
                            for layout in result.get('layouts_used', [])
                        ]
                    }
                    output.append(output_item)
                
                with open('results.json', 'w', encoding='utf-8') as f:
                    json.dump(output, f, indent=4, ensure_ascii=False)
                
        except Exception as e:
            logging.error(f"Erro ao exportar resultados JSON: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())
            raise

def main():
    """Função principal do otimizador"""
    try:
        # Carrega dados de entrada
        with open('dados_entrada.json', 'r', encoding='utf-8') as f:
            input_data = json.load(f)
        
        logging.info("Iniciando otimização de produção")
        optimizer = LayoutOptimizer(input_data)
        
        # Processa cada peça como uma ordem separada
        results = {}
        for i, piece in enumerate(optimizer.pieces, 1):
            order_id = f'order_{i}'
            logging.info(f"Processando {piece['pattern']} (Ordem {order_id})")
            
            # Cria estrutura de ordem
            order = {
                'id': order_id,
                'pieces': [piece],
                'fabric_width': optimizer.fabrics[piece['fabrics'][0]]['fabric_width'],
                'max_layers': optimizer.fabrics[piece['fabrics'][0]]['max_layers'],
                'max_length': optimizer.config['max_total_length']
            }
            
            # Otimiza ordem
            result = optimizer.optimize_order(order)
            results[order_id] = result
            
            # Log dos resultados
            if result:
                logging.info(f"Solução encontrada para {piece['pattern']}")
                logging.info("Produção por tamanho:")
                for size in optimizer.sizes:
                    prod = result['production'][size]
                    over = result['overproduction'][size]
                    logging.info(f"  {size}: {prod} (Excesso: {over})")
                    
                metrics = result['metrics']
                logging.info(f"Custo total: R$ {metrics['total_cost']:.2f}")
                logging.info(f"Metros de tecido: {metrics['fabric_meters']:.2f} m")
                logging.info(f"Desperdício: {metrics['fabric_waste']:.4f} m²")
            else:
                logging.warning(f"Não foi possível encontrar solução para {piece['pattern']}")
        
        # Exporta resultados
        if any(results.values()):
            optimizer.export_results(results)
            optimizer.export_results_json(results)
            logging.info("Resultados exportados para results.xlsx e results.json")
        else:
            logging.warning("Nenhuma solução encontrada")
            
    except Exception as e:
        logging.error(f"Erro na execução: {str(e)}")
        raise

if __name__ == "__main__":
    main()