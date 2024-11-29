import json
import logging
from ortools.linear_solver import pywraplp
from openpyxl import Workbook
import unicodedata
import traceback

# Logging Configuration
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("debug.log"), logging.StreamHandler()],
)


class UnitConverter:
    """Classe utilitária para conversão de unidades e cálculos de métricas"""
    
    @staticmethod
    def mm_to_m(value):
        """Converte milímetros para metros"""
        return value / 1000.0
    
    @staticmethod
    def cm2_to_m2(value):
        """Converte centímetros quadrados para metros quadrados"""
        return value / 10000.0
    
    @staticmethod
    def calculate_layout_metrics(layout, num_layers):
        """Calcula todas as métricas do layout com as unidades corretas"""
        return {
            'length_meters': UnitConverter.mm_to_m(layout['layout_length']),
            'waste_area_m2': UnitConverter.cm2_to_m2(layout['waste_area']),
            'total_area_m2': UnitConverter.cm2_to_m2(layout['total_area']),
            'perimeter_meters': UnitConverter.mm_to_m(layout['total_perimeter']),
            'fabric_width_m': UnitConverter.mm_to_m(layout['fabric_width']),
            'utilization': layout['utilization'],
            'num_layers': num_layers
        }
    
    @staticmethod
    def calculate_waste_metrics(layout):
        """Calcula métricas de desperdício do layout"""
        try:
            # Converte áreas de mm² para m²
            total_area_m2 = layout['total_area'] / 10000  # mm² para m²
            waste_area_m2 = layout['waste_area'] / 10000  # mm² para m²
            fabric_width_m = layout['fabric_width'] / 1000  # mm para m
            
            # Calcula desperdício em metros lineares
            waste_meters = waste_area_m2 / fabric_width_m if fabric_width_m > 0 else 0
            
            # Calcula percentual de desperdício
            waste_percentage = (layout['waste_area'] / layout['total_area']) * 100 if layout['total_area'] > 0 else 0
            
            return {
                'fabric_waste_area': waste_area_m2,
                'fabric_waste_meters': waste_meters,
                'total_waste_percentage': waste_percentage
            }
            
        except Exception as e:
            logging.error(f"Erro no cálculo de métricas de desperdício: {str(e)}")
            raise

class LayoutOptimizer:
    def __init__(self, input_data):
        self.config = input_data['general_configuration']
        self.layouts = input_data.get('layout', input_data.get('layouts', []))
        self.fabrics = {f['fabric']: f for f in input_data['fabrics']}
        self.pieces = input_data['pieces']
        
        # Extrair tamanhos de todas as fontes possíveis
        sizes_from_pieces = set()
        sizes_from_layouts = set()
        
        # Extrair tamanhos das peças (demanda)
        for piece in self.pieces:
            if 'quantity' in piece:
                sizes_from_pieces.update(piece['quantity'].keys())
        
        # Extrair tamanhos dos layouts
        for layout in self.layouts:
            for piece in layout['pieces']:
                if 'size_grade' in piece:
                    sizes_from_layouts.update(piece['size_grade'].keys())
        
        # Combinar todos os tamanhos encontrados
        self.sizes = sorted(sizes_from_pieces.union(sizes_from_layouts))
        
        if not self.sizes:
            raise ValueError("Não foi possível extrair os tamanhos dos dados de entrada")
        
        logging.info(f"Tamanhos extraídos automaticamente: {self.sizes}")
        
        # Restante da inicialização...
        self.overproduction_penalty = self.config.get('overproduction_percentage', 0.05)
        self.unit_waste_cost = self.config.get('waste_cost', 0.1)
        self.optimality_gap = self.config.get('optimality_gap', 0.01)
        self.max_memory_mb = self.config.get('max_memory_mb', 2048)
        
        # Novos parâmetros
        self.layout_change_penalty = self.config.get('layout_change_penalty', 1000)
        self.min_layers_per_layout = self.config.get('min_layers_per_layout', 5)
        self.waste_penalty_factor = self.config.get('waste_penalty_factor', 2.0)
        
        # Log dos parâmetros carregados
        logging.info("Parâmetros de otimização:")
        logging.info(f"  Layout change penalty: {self.layout_change_penalty}")
        logging.info(f"  Min layers per layout: {self.min_layers_per_layout}")
        logging.info(f"  Waste penalty factor: {self.waste_penalty_factor}")

        # Validação dos novos parâmetros
        if self.layout_change_penalty < 0:
            raise ValueError("layout_change_penalty deve ser não-negativo")
        
        if self.min_layers_per_layout < 1:
            raise ValueError("min_layers_per_layout deve ser pelo menos 1")
        
        if self.waste_penalty_factor < 0:
            raise ValueError("waste_penalty_factor deve ser não-negativo")
        
        for fabric in self.fabrics.values():
            fabric['price_per_square_meter'] = (
                fabric['price_per_linear_meter'] / (fabric['fabric_width'] / 1000)
            )
            logging.info(f"Preço por m² do tecido {fabric['fabric']}: R${fabric['price_per_square_meter']:.2f}")
       




    def optimize_production(self):
        results = {}
        
        for i, piece in enumerate(self.pieces, 1):
            order_id = f'order_{i}'
            pattern = piece['pattern']
            logging.info(f"Processando {pattern} (Ordem {order_id})")
            
            # Cria estrutura de ordem
            order = {
                'id': order_id,
                'pattern': pattern,
                'pieces': [piece],
                'fabric_width': self.fabrics[piece['fabrics'][0]]['fabric_width'],
                'max_layers': self.fabrics[piece['fabrics'][0]]['max_layers'],
                'max_length': self.config['max_total_length']
            }
            
            # Otimiza ordem
            result = self.optimize_order(order)
            if result:
                # Garante que todas as métricas existam com valores padrão
                metrics = {
                    'fabric_meters': 0.0,
                    'fabric_cost': 0.0,
                    'cutting_cost': 0.0,
                    'layout_cost': 0.0,
                    'layer_cost': 0.0,
                    'waste_cost': 0.0,
                    'total_cost': 0.0,
                    'fabric_waste_area': 0.0,
                    'fabric_waste_meters': 0.0,
                    'total_waste_percentage': 0.0
                }
                
                # Atualiza com os valores calculados
                metrics.update(result['metrics'])
                
                # Recalcula o custo total para garantir consistência
                metrics['total_cost'] = sum([
                    metrics['fabric_cost'],
                    metrics['cutting_cost'],
                    metrics['layout_cost'],
                    metrics['layer_cost'],
                    metrics['waste_cost']
                ])
                
                result['metrics'] = metrics
                results[order_id] = result
                
                # Validação final das métricas
                self._validate_metrics(result['metrics'])
                
        return results

    def _validate_layout_metrics(self, layout_data):
        """Valida as métricas dos layouts"""
        # Comprimento em metros deve ser maior que 0
        if layout_data['length_meters'] <= 0:
            raise ValueError(f"Comprimento inválido: {layout_data['length_meters']}")
            
        # Comprimento em metros deve ser consistente com o valor em milímetros
        expected_length = layout_data['layout_length'] / 1000
        if abs(layout_data['length_meters'] - expected_length) > 0.001:
            raise ValueError(
                f"Inconsistência no comprimento: "
                f"esperado {expected_length}m, "
                f"obtido {layout_data['length_meters']}m"
            )
        

    def _validate_waste_metrics(self, layout_data):
        """Valida métricas de desperdício do layout"""
        try:
            # Converte unidades
            metrics = UnitConverter.calculate_layout_metrics(layout_data, 1)
            waste_metrics = UnitConverter.calculate_waste_metrics(metrics)
            
            # Valida utilização
            calculated_utilization = 1 - (waste_metrics['fabric_waste_area'] / metrics['total_area_m2'])
            if abs(calculated_utilization - layout_data['utilization']) > 0.01:
                raise ValueError(
                    f"Inconsistência no cálculo de utilização do layout {layout_data['id']}: "
                    f"calculado {calculated_utilization:.4f}, informado {layout_data['utilization']:.4f}"
                )
                
            return waste_metrics
        
        except Exception as e:
            logging.error(f"Erro na validação de métricas de desperdício: {str(e)}")
            raise
    
    def _validate_metrics(self, metrics):
        """Valida a consistência das métricas calculadas"""
        try:
            # Verifica valores negativos
            for key, value in metrics.items():
                if value < 0:
                    raise ValueError(f"Métrica {key} com valor negativo: {value}")
            
            # Verifica consistência do custo total
            expected_total = sum([
                metrics['fabric_cost'],
                metrics['cutting_cost'],
                metrics['layout_cost'],
                metrics['layer_cost'],
                metrics['waste_cost']
            ])
            
            if abs(expected_total - metrics['total_cost']) > 0.01:
                raise ValueError(
                    f"Inconsistência no custo total: "
                    f"esperado R${expected_total:.2f}, "
                    f"obtido R${metrics['total_cost']:.2f}"
                )
            
                
            # Verifica consistência do desperdício
            if metrics['fabric_waste_meters'] > metrics['fabric_meters']:
                raise ValueError("Metros desperdiçados maior que metros totais")
            
            # Validação específica para métricas de desperdício
            if metrics['fabric_waste_area'] > 0:
                expected_waste_meters = metrics['fabric_waste_area'] / (self.fabrics[self.pieces[0]['fabrics'][0]]['fabric_width'] / 1000)
                if abs(metrics['fabric_waste_meters'] - expected_waste_meters) > 0.0001:
                    raise ValueError(
                        f"Inconsistência nos metros desperdiçados: "
                        f"esperado {expected_waste_meters:.6f}m, "
                        f"obtido {metrics['fabric_waste_meters']:.6f}m"
                    )
                
            if metrics['fabric_meters'] <= 0:
                raise ValueError("Metros de tecido deve ser maior que zero")
                
            if metrics['total_cost'] <= 0:
                raise ValueError("Custo total deve ser maior que zero")
                
            # Novas validações para desperdício
            if metrics['fabric_waste_area'] < 0:
                raise ValueError("Área desperdiçada não pode ser negativa")
                
            if metrics['fabric_waste_meters'] < 0:
                raise ValueError("Metros desperdiçados não pode ser negativo")
                
            if not (0 <= metrics['total_waste_percentage'] <= 100):
                raise ValueError(
                    f"Percentual de desperdício deve estar entre 0 e 100, encontrado: "
                    f"{metrics['total_waste_percentage']:.2f}%"
                )
                
        except Exception as e:
            logging.error(f"Erro na validação de métricas: {str(e)}")
            raise
    
    def _validate_costs(self, layout_data):
        try:
            fabric = self.fabrics[layout_data['fabric']]
            tolerance = 0.01
            num_layers = layout_data['num_layers']
            length_meters = layout_data['layout_length'] / 1000  # mm para m
            
            # Conversões para cálculo do desperdício
            waste_area_m2 = layout_data['waste_area'] / 1_000_000  # mm² para m²
            fabric_width_m = layout_data['fabric_width'] / 1000    # mm para m
            waste_meters = waste_area_m2 / fabric_width_m
            
            # Custos esperados
            expected_costs = {
                'fabric_cost': length_meters * fabric['price_per_linear_meter'] * num_layers,
                'cutting_cost': (layout_data['total_perimeter'] / 1000) * fabric['cost_per_cut_meter'] * num_layers,
                'layout_cost': length_meters * fabric['cost_per_layout_meter'] * num_layers,
                'layer_cost': fabric['cost_per_layer'] * num_layers,
                'waste_cost': waste_meters * fabric['price_per_linear_meter'] * num_layers
            }
            
            # Validação
            for cost_type, expected in expected_costs.items():
                actual = layout_data['costs'].get(cost_type, 0)
                if abs(actual - expected) > tolerance:
                    raise ValueError(
                        f"Erro no cálculo do {cost_type}: "
                        f"esperado R${expected:.2f}, "
                        f"obtido R${actual:.2f}"
                    )
                    
        except Exception as e:
            logging.error(f"Erro na validação dos custos: {str(e)}")
            raise



    def _calculate_waste_cost(self, layout, num_layers):
        try:
            fabric = self.fabrics[layout['fabric']]
            
            # Conversões corretas de unidades
            waste_area_m2 = layout['waste_area'] / 1_000_000  # mm² para m²
            fabric_width_m = layout['fabric_width'] / 1000    # mm para m
            waste_meters = waste_area_m2 / fabric_width_m
            
            # Calcula o custo do desperdício
            waste_cost = waste_meters * fabric['price_per_linear_meter'] * num_layers
            
            # Validações
            if waste_cost < 0:
                raise ValueError(f"Custo de desperdício negativo: {waste_cost}")
            if waste_meters > layout['layout_length'] / 1000:
                raise ValueError(f"Metros desperdiçados ({waste_meters}) maior que comprimento do layout ({layout['layout_length']/1000})")
                
            return {
                'waste_cost': waste_cost,
                'fabric_waste_area': waste_area_m2,
                'fabric_waste_meters': waste_meters
            }
        except Exception as e:
            logging.error(f"Erro no cálculo do custo de desperdício: {str(e)}")
            raise


    # def calculate_layout_costs(self, layout, num_layers, fabric_cost):
    #     """
    #     Calcula custos do layout separando claramente cada componente
        
    #     Args:
    #         layout (dict): Informações do layout
    #         num_layers (int): Número de camadas
    #         fabric_cost (float): Custo do tecido por metro linear
        
    #     Returns:
    #         dict: Custos calculados separadamente
    #     """
    #     try:
    #         fabric = self.fabrics[layout['fabric']]
            
    #         # 1. Custo do tecido
    #         fabric_cost = (
    #             fabric['price_per_linear_meter'] * 
    #             layout['layout_length'] * 
    #             num_layers / 1000
    #         )
            
    #         # 2. Custo de corte
    #         cutting_cost = (
    #             fabric['cost_per_cut_meter'] * 
    #             layout['total_perimeter'] * 
    #             num_layers / 1000
    #         )
            
    #         # 3. Custo de setup do layout (sem incluir custo por camada)
    #         layout_cost = (
    #             fabric['cost_per_layout_meter'] * 
    #             layout['layout_length'] * 
    #             num_layers / 1000
    #         )
            
    #         # 4. Custo de desperdício
    #         waste_cost = self._calculate_waste_cost(layout, num_layers)
            
    #         costs = {
    #             'fabric_cost': fabric_cost,
    #             'cutting_cost': cutting_cost,
    #             'layout_cost': layout_cost,
    #             'waste_cost': waste_cost
    #         }
            
    #         # Valida os custos calculados
    #         self._validate_costs(costs, layout, num_layers)
            
    #         return costs
        
    #     except Exception as e:
    #         logging.error(f"Erro ao calcular custos do layout: {str(e)}")
    #         raise

    def calculate_layout_costs(self, layout, num_layers):
        """
        Calcula os custos do layout considerando o número de camadas.
        Todas as medidas de entrada estão em mm e mm² e são convertidas para m e m².
        """
        try:
            fabric = self.fabrics[layout['fabric']]
            
            # Conversão de unidades (mm para m)
            layout_length_m = layout['layout_length'] / 1000  # mm para m
            waste_area_m2 = layout['waste_area'] / 1_000_000  # mm² para m²
            fabric_width_m = layout['fabric_width'] / 1000    # mm para m
            perimeter_m = layout['total_perimeter'] / 1000    # mm para m
            
            # Cálculo do desperdício em metros lineares
            waste_meters = waste_area_m2 / fabric_width_m
            
            # Cálculo dos custos
            costs = {
                'fabric_cost': layout_length_m * fabric['price_per_linear_meter'] * num_layers,
                'cutting_cost': perimeter_m * fabric['cost_per_cut_meter'] * num_layers,
                'layout_cost': layout_length_m * fabric['cost_per_layout_meter'] * num_layers,
                'layer_cost': fabric['cost_per_layer'] * num_layers,
                'waste_cost': waste_meters * fabric['price_per_linear_meter'] * num_layers
            }
            
            # Cálculo do custo total
            costs['total_cost'] = sum(costs.values())
            
            # Métricas não monetárias para análise
            costs['fabric_meters'] = layout_length_m * num_layers
            costs['fabric_waste_area'] = waste_area_m2 * num_layers
            costs['fabric_waste_meters'] = waste_meters * num_layers
            costs['total_waste_percentage'] = (waste_area_m2 / (layout_length_m * fabric_width_m)) * 100
            
            logging.debug(f"""
                Cálculo de custos para layout {layout['id']}:
                - Comprimento: {layout_length_m:.3f} m
                - Largura: {fabric_width_m:.3f} m
                - Área desperdiçada: {waste_area_m2:.6f} m²
                - Metros desperdiçados: {waste_meters:.6f} m
                - Número de camadas: {num_layers}
                - Custo do desperdício: R${costs['waste_cost']:.2f}
            """)
            
            return costs
            
        except Exception as e:
            logging.error(f"Erro no cálculo de custos do layout {layout.get('id')}: {str(e)}")
            logging.error(f"Layout: {layout}")
            raise

    def _calculate_layout_metrics(self, layout_data):
        """Calcula métricas para um layout específico"""
        try:
            fabric = self.fabrics[layout_data['fabric']]
            num_layers = layout_data['num_layers']
            
            # Cálculos básicos
            length_meters = layout_data['length_meters'] / 1000  # mm para m
            fabric_cost = length_meters * fabric['price_per_linear_meter'] * num_layers
            cutting_cost = length_meters * fabric['cost_per_cut_meter'] * num_layers
            layout_cost = length_meters * fabric['cost_per_layout_meter']
            layer_cost = fabric['cost_per_layer'] * num_layers
            
            # Cálculo do desperdício
            waste_metrics = self._calculate_waste_cost(layout_data, num_layers)
            
            return {
                'fabric_cost': fabric_cost,
                'cutting_cost': cutting_cost,
                'layout_cost': layout_cost,
                'layer_cost': layer_cost,
                'waste_cost': waste_metrics['waste_cost'],  # ✅ Incluído aqui
                'fabric_waste_area': waste_metrics['fabric_waste_area'],
                'fabric_waste_meters': waste_metrics['fabric_waste_meters']
            }
        except Exception as e:
            logging.error(f"Erro no cálculo das métricas do layout: {str(e)}")
            raise

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
    



    
    # def _process_solution(self, solver, layouts, x, overproduction, demand, pattern, piece, order):
    #     """
    #     Processa a solução do solver e calcula todas as métricas
        
    #     Args:
    #         solver: Solver do OR-Tools
    #         layouts: Lista de layouts disponíveis
    #         x: Variáveis de decisão (número de camadas por layout)
    #         overproduction: Variáveis de superprodução
    #         demand: Demanda por tamanho
    #         pattern: Padrão sendo processado
    #         piece: Informações da peça
    #         order: Informações da ordem
        
    #     Returns:
    #         dict: Resultado processado com todas as métricas
    #     """
    #     try:
    #         result = {
    #             'status': 'optimal',
    #             'pattern': piece['pattern'],
    #             'demand': piece['quantity'],
    #             'production': {},
    #             'overproduction': {},
    #             'layouts_used': [],
    #             'metrics': {
    #                 'fabric_meters': 0,
    #                 'fabric_cost': 0,
    #                 'cutting_cost': 0,
    #                 'layout_cost': 0,
    #                 'layer_cost': 0,
    #                 'waste_cost': 0,
    #                 'total_cost': 0,
    #                 'fabric_waste_area': 0,
    #                 'fabric_waste_meters': 0
    #             }
    #         }

    #         total_production = {size: 0 for size in self.sizes}
    #         total_layers = 0

    #         # Processa layouts utilizados
    #         for layout in layouts:
    #             num_layers = int(x[layout['id']].solution_value())
    #             if num_layers > 0:
    #                 fabric = self.fabrics[piece['fabrics'][0]]
    #                 total_layers += num_layers
                    
    #                 # Calcula produção por tamanho
    #                 production_per_size = {}
    #                 for size in self.sizes:
    #                     if size in layout['pieces'][0]['size_grade']:
    #                         qty_per_layer = layout['pieces'][0]['size_grade'][size]
    #                         production_per_size[size] = qty_per_layer * num_layers
    #                         total_production[size] += production_per_size[size]

    #                 # Calcula custos
    #                 costs = self.calculate_layout_costs(layout, num_layers, fabric['price_per_linear_meter'])
                    
    #                 # Adiciona layout usado (corrigido para usar num_layers)
    #                 layout_info = {
    #                     'layout_id': layout['id'],
    #                     'num_layers': num_layers,  # Alterado de 'layers' para 'num_layers'
    #                     'length_meters': layout['layout_length'] / 1000,
    #                     'utilization': layout['utilization'],
    #                     'waste_area': layout['waste_area'] / 1_000_000,
    #                     'production_per_size': production_per_size,
    #                     'costs': costs
    #                 }
    #                 result['layouts_used'].append(layout_info)

    #                 # Atualiza métricas
    #                 result['metrics']['fabric_meters'] += layout['layout_length'] * num_layers / 1000
    #                 result['metrics']['fabric_cost'] += costs['fabric_cost']
    #                 result['metrics']['cutting_cost'] += costs['cutting_cost']
    #                 result['metrics']['layout_cost'] += costs['layout_cost']
    #                 result['metrics']['waste_cost'] += costs['waste_cost']
    #                 result['metrics']['fabric_waste_area'] += layout['waste_area'] * num_layers / 1_000_000
    #                 result['metrics']['fabric_waste_meters'] += (layout['waste_area'] / layout['fabric_width']) * num_layers / 1000

    #         # Calcula custo por camada uma única vez
    #         fabric = self.fabrics[piece['fabrics'][0]]
    #         result['metrics']['layer_cost'] = fabric['cost_per_layer'] * total_layers

    #         # Atualiza produção e superprodução
    #         result['production'] = total_production
    #         result['overproduction'] = {
    #             size: max(0, total_production[size] - piece['quantity'][size])
    #             for size in self.sizes if size in piece['quantity']
    #         }

    #         # Calcula custo total
    #         result['metrics']['total_cost'] = (
    #             result['metrics']['fabric_cost'] +
    #             result['metrics']['cutting_cost'] +
    #             result['metrics']['layout_cost'] +
    #             result['metrics']['layer_cost'] +
    #             result['metrics']['waste_cost']
    #         )

    #         # Valida resultado
    #         self._validate_solution(result, piece['quantity'])

    #         return result

    #     except Exception as e:
    #         logging.error(f"Erro ao processar solução: {str(e)}")
    #         raise


    def calculate_layout_costs(self, layout, num_layers):
        """Calcula os custos do layout"""
        try:
            fabric = self.fabrics[layout['fabric']]
            
            # Primeiro calcula métricas básicas do layout
            base_metrics = UnitConverter.calculate_layout_metrics(layout, num_layers)
            
            # Prepara dados para cálculo de desperdício
            waste_data = {
                'total_area': layout['total_area'],  # área original em mm²
                'waste_area': layout['waste_area'],  # área de desperdício em mm²
                'fabric_width': layout['fabric_width']  # largura em mm
            }
            
            # Calcula métricas de desperdício
            waste_metrics = UnitConverter.calculate_waste_metrics(waste_data)
            
            # Calcula custos
            costs = {
                'fabric_meters': base_metrics['length_meters'] * num_layers,
                'fabric_cost': base_metrics['length_meters'] * fabric['price_per_linear_meter'] * num_layers,
                'cutting_cost': base_metrics['perimeter_meters'] * fabric['cost_per_cut_meter'] * num_layers,
                'layout_cost': base_metrics['length_meters'] * fabric['cost_per_layout_meter'] * num_layers,
                'layer_cost': fabric['cost_per_layer'] * num_layers,
                'waste_cost': waste_metrics['fabric_waste_meters'] * fabric['price_per_linear_meter'] * num_layers
            }
            
            # Adiciona métricas de desperdício aos custos
            costs.update({
                'fabric_waste_area': waste_metrics['fabric_waste_area'] * num_layers,
                'fabric_waste_meters': waste_metrics['fabric_waste_meters'] * num_layers,
                'total_waste_percentage': waste_metrics['total_waste_percentage']
            })
            
            return costs
            
        except Exception as e:
            logging.error(f"Erro no cálculo de custos do layout {layout.get('id')}: {str(e)}")
            raise


    def _process_solution(self, solver, x, y, overproduction, filtered_layouts, demand_quantity, pattern):
        """Processa a solução do solver e calcula todas as métricas"""
        try:
            # Captura os valores das variáveis imediatamente após a solução
            layout_solutions = {
                layout['id']: x[layout['id']].solution_value() 
                for layout in filtered_layouts
            }
            
            result = {
                'pattern': pattern,
                'metrics': {
                    'fabric_meters': 0.0,
                    'fabric_cost': 0.0,
                    'cutting_cost': 0.0,
                    'layout_cost': 0.0,
                    'layer_cost': 0.0,
                    'waste_cost': 0.0,
                    'total_cost': 0.0,
                    'fabric_waste_area': 0.0,
                    'fabric_waste_meters': 0.0,
                    'total_waste_percentage': 0.0
                },
                'demand': {size: demand_quantity.get(size, 0) for size in self.sizes},  # Modificado
                'production': {size: 0 for size in self.sizes},  # Modificado
                'overproduction': {size: 0 for size in self.sizes},  # Modificado
                'layouts_used': []
            }

            total_area = 0
            total_waste_area = 0

            # Processa cada layout utilizado
            for layout in filtered_layouts:
                num_layers = int(layout_solutions[layout['id']])
                if num_layers > 0:
                    # Calcula métricas básicas do layout
                    metrics = UnitConverter.calculate_layout_metrics(layout, num_layers)
                    
                    # Calcula métricas de desperdício
                    waste_metrics = UnitConverter.calculate_waste_metrics({
                        'total_area': layout['total_area'],
                        'waste_area': layout['waste_area'],
                        'fabric_width': layout['fabric_width']
                    })
                    
                    # Calcula custos do layout
                    layout_costs = self.calculate_layout_costs(layout, num_layers)
                    
                    # Prepara informações do layout
                    layout_info = {
                        'layout_id': layout['id'],
                        'num_layers': num_layers,
                        'length_meters': metrics['length_meters'],
                        'utilization': layout['utilization'],
                        'waste_area': waste_metrics['fabric_waste_area'],
                        'production_per_size': {size: 0 for size in self.sizes},  # Modificado
                        'costs': layout_costs
                    }
                    
                    # Calcula produção por tamanho
                    for size in self.sizes:
                        if size in layout['pieces'][0]['size_grade']:
                            qty = layout['pieces'][0]['size_grade'][size] * num_layers
                            layout_info['production_per_size'][size] = qty
                            result['production'][size] = result['production'].get(size, 0) + qty
                    
                    # Atualiza métricas totais
                    result['metrics']['fabric_meters'] += metrics['length_meters'] * num_layers
                    result['metrics']['fabric_cost'] += layout_costs['fabric_cost']
                    result['metrics']['cutting_cost'] += layout_costs['cutting_cost']
                    result['metrics']['layout_cost'] += layout_costs['layout_cost']
                    result['metrics']['layer_cost'] += layout_costs['layer_cost']
                    result['metrics']['waste_cost'] += layout_costs['waste_cost']
                    result['metrics']['fabric_waste_area'] += waste_metrics['fabric_waste_area'] * num_layers
                    result['metrics']['fabric_waste_meters'] += waste_metrics['fabric_waste_meters'] * num_layers
                    
                    # Acumula áreas para cálculo do desperdício total
                    total_area += layout['total_area'] * num_layers
                    total_waste_area += layout['waste_area'] * num_layers
                    
                    result['layouts_used'].append(layout_info)
            
            # Calcula percentual total de desperdício
            if total_area > 0:
                result['metrics']['total_waste_percentage'] = (total_waste_area / total_area) * 100
            
            # Calcula custo total
            result['metrics']['total_cost'] = sum([
                result['metrics']['fabric_cost'],
                result['metrics']['cutting_cost'],
                result['metrics']['layout_cost'],
                result['metrics']['layer_cost'],
                result['metrics']['waste_cost']
            ])
            
            # Calcula superprodução final
            for size in self.sizes:
                if size in demand_quantity:
                    result['overproduction'][size] = max(0, result['production'][size] - demand_quantity[size])
            
            return result
            
        except Exception as e:
            logging.error(f"Erro no processamento da solução: {str(e)}")
            logging.error(traceback.format_exc())
            raise

    def _validate_final_results(self, result):
        """Validação final dos resultados"""
        try:
            # Valida produção vs demanda
            for size in self.sizes:
                if result['production'][size] < result['demand'][size]:
                    raise ValueError(f"Produção insuficiente para tamanho {size}")
                if result['production'][size] > result['demand'][size] * 1.05:
                    raise ValueError(f"Superprodução excessiva para tamanho {size}")
                    
            # Valida custos totais
            total_cost = sum([
                result['metrics']['fabric_cost'],
                result['metrics']['cutting_cost'],
                result['metrics']['layout_cost'],
                result['metrics']['layer_cost'],
                result['metrics']['waste_cost']
            ])
            
            if abs(total_cost - result['metrics']['total_cost']) > 0.01:
                raise ValueError(
                    f"Inconsistência no custo total: "
                    f"calculado {total_cost:.2f}, "
                    f"informado {result['metrics']['total_cost']:.2f}"
                )
                
        except Exception as e:
            logging.error(f"Erro na validação final: {str(e)}")
            raise

    def _validate_units(self, result):
        """Valida as unidades das métricas calculadas"""
        try:
            # Validações existentes
            if result['metrics']['fabric_meters'] < 0:
                raise ValueError("Metros de tecido não pode ser negativo")
            
            # Novas validações
            for layout in result['layouts_used']:
                # Valida conversões
                length_m = UnitConverter.mm_to_m(layout['layout_length'])
                waste_m2 = UnitConverter.cm2_to_m2(layout['waste_area'])
                
                # Compara com valores calculados
                if abs(layout['length_meters'] - length_m) > 0.001:
                    raise ValueError(f"Inconsistência na conversão de comprimento no layout {layout['layout_id']}")
                    
                if abs(layout['waste_area'] - waste_m2) > 0.001:
                    raise ValueError(f"Inconsistência na conversão de área no layout {layout['layout_id']}")
                    
        except Exception as e:
            logging.error(f"Erro na validação de unidades: {str(e)}")
            raise
    
    def optimize_order(self, order):
        """
        Optimize an specific order allowing multiple layouts
        
        Calculation formulas:

            1. Fabric cost:
        fabric_cost = price_per_linear_meter * layout_length * num_layers / 1000

            2. Cutting cost:
        cutting_cost = cost_per_cut_meter * total_perimeter * num_layers / 1000

            3. Setup cost:
        setup_cost = cost_per_layer * num_layers + 
                        cost_per_layout_meter * layout_length * num_layers / 1000

            4. Waste cost:
        waste_cost = (waste_area_m2 * fabric_price_per_m2 * num_layers * waste_penalty_factor)

            5. Total cost:
        total_cost = fabric_cost + cutting_cost + setup_cost + layer_cost + waste_cost
        
        """
        try:
            logging.info("=== Iniciando Otimização ===")
            logging.info(f"Demanda por tamanho: {order['pieces'][0]['quantity']}")
            self._validate_input_data(order)

            piece = order['pieces'][0]
            pattern = order['pattern']
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

            # Valida tamanhos em todos os layouts filtrados
            for layout in filtered_layouts:
                self._validate_layout_sizes(layout)
            
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
            # Inicializa variáveis para cada layout
            for layout in filtered_layouts:
                x[layout['id']] = solver.IntVar(0, order['max_layers'], f'x_{layout["id"]}')
                y[layout['id']] = solver.IntVar(0, 1, f'y_{layout["id"]}')
                
                # Relaciona x e y: se y=0, x deve ser 0
                solver.Add(x[layout['id']] <= order['max_layers'] * y[layout['id']])
            
            # Garante uso de pelo menos um layout
            solver.Add(solver.Sum(y[l['id']] for l in filtered_layouts) >= 1)
            
            
            # Cria variáveis para superprodução
            overproduction = {}
            for size in self.sizes:
                overproduction[size] = solver.NumVar(0, solver.infinity(), f'over_{size}')
            
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
                    solver.Add(total_production <= demand_quantity[size] * (1 + self.overproduction_penalty))  # 5% máximo
                    
                    logging.info(f"Configurada restrição para tamanho {size}:")
                    logging.info(f"  Demanda: {demand_quantity[size]}")
                    logging.info(f"  Produção mínima: {demand_quantity[size]}")
                    logging.info(f"  Produção máxima: {demand_quantity[size] * 1.05}")
            
            # Função objetivo: minimizar custos totais
            # objective = solver.Sum([
            #     x[layout['id']] * (
            #         # Custo do tecido
            #         # layout['layout_length'] * self.fabrics[fabric]['price_per_linear_meter'] / 1000 +
            #         # Custo de corte
            #         layout['total_perimeter'] * self.fabrics[fabric]['cost_per_cut_meter'] / 1000 +
            #         # Custo por camada
            #         self.fabrics[fabric]['cost_per_layer'] +
            #         # Penalidade por desperdício
            #         layout['waste_area'] * self.unit_waste_cost / 1_000_000
            #     )
            #     for layout in filtered_layouts
            # ])
            
            # solver.Minimize(objective)
   

            # Função objetivo atualizada: minimizar desperdício e custos reais + penalidade por superprodução
            # objective = (
            #     solver.Sum([
            #         x[layout['id']] * (
            #             # Custos base
            #             layout['layout_length'] * self.fabrics[layout['fabric']]['price_per_linear_meter'] / 1000 +
            #             layout['total_perimeter'] * self.fabrics[layout['fabric']]['cost_per_cut_meter'] / 1000 +
            #             self.fabrics[layout['fabric']]['cost_per_layer'] +
            #             self.fabrics[layout['fabric']]['cost_per_layout_meter'] * layout['layout_length'] / 1000 +
                        
            #             # Penalidade por desperdício
            #             (layout['waste_area'] / layout['total_area']) * (
            #                 self.waste_penalty_factor * 
            #                 layout['layout_length'] * 
            #                 self.fabrics[layout['fabric']]['price_per_linear_meter'] / 1000
            #             )
            #         )
            #         for layout in filtered_layouts
            #     ]) +
            #     # Penalidade por múltiplos layouts 
            #     self.layout_change_penalty * solver.Sum([y[l['id']] for l in filtered_layouts]) +
            #     # Penalidade por superprodução 
            #     solver.Sum([
            #         overproduction[size] * self.fabrics[fabric]['price_per_linear_meter'] * self.waste_penalty_factor
            #         for size in self.sizes
            #     ])
            # )

             # Função objetivo corrigida
            objective = solver.Sum([
                x[layout['id']] * UnitConverter.cm2_to_m2(float(layout['waste_area']))  +  # Converte para float e m²
                solver.Sum([
                    overproduction[size] * 0.01 * float(layout['waste_area']) / float(layout['total_area'])
                    for size in self.sizes
                ])
                for layout in filtered_layouts
            ])

            # Adiciona restrição de número mínimo de camadas
            for layout in filtered_layouts:
                solver.Add(
                    x[layout['id']] >= self.min_layers_per_layout * y[layout['id']]
                )

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
                    # Processa solução - Corrigindo a chamada
                    result = self._process_solution(
                        solver=solver,
                        x=x,
                        y=y,
                        overproduction=overproduction,
                        filtered_layouts=filtered_layouts, 
                        demand_quantity=demand_quantity,
                        pattern=pattern
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
                    self._validate_solution(result, demand_quantity)

                     
                    logging.info("=== Resultados da Otimização ===")
                    logging.info(f"Status: {status}")
                    logging.info(f"Valor objetivo: {solver.Objective().Value()}")
                    logging.info("Custos detalhados:")
                    for metric, value in result['metrics'].items():
                        logging.info(f"  {metric}: {value:.2f}")
                    
                    return result
                
            return None
            
        except Exception as e:
            logging.error(f"Erro na otimização: {str(e)}")
            import traceback
            logging.error(traceback.format_exc())
            return None
    
    def _validate_input_data(self, order):
        """Valida dados de entrada"""
        if not order.get('pieces'):
            raise ValueError("Ordem deve conter peças")
            
        piece = order['pieces'][0]
        if not all(key in piece for key in ['pattern', 'quantity', 'fabrics']):
            raise ValueError("Dados da peça incompletos")
            
        # Validar se todos os tamanhos da demanda estão nos tamanhos extraídos
        demand_sizes = set(piece['quantity'].keys())
        if not demand_sizes.issubset(set(self.sizes)):
            invalid_sizes = demand_sizes - set(self.sizes)
            raise ValueError(f"Tamanhos inválidos na demanda: {invalid_sizes}")
    
    def _validate_layout_sizes(self, layout):
        """Valida se os tamanhos no layout são compatíveis com os tamanhos extraídos"""
        for piece in layout['pieces']:
            layout_sizes = set(piece['size_grade'].keys())
            if not layout_sizes.issubset(set(self.sizes)):
                invalid_sizes = layout_sizes - set(self.sizes)
                raise ValueError(f"Tamanhos inválidos no layout {layout['id']}: {invalid_sizes}")

    def _validate_solution(self, solution, demand):
        """Valida resultado da otimização"""
        if not solution:  # ✅ Verifica se solution não é None
            raise ValueError("Solução inválida (None)")
            
        # Verifica se todas as métricas necessárias existem
        required_metrics = [
            'fabric_meters', 'fabric_cost', 'cutting_cost', 
            'layout_cost', 'layer_cost', 'waste_cost',
            'total_cost'
        ]
        
        for metric in required_metrics:
            if metric not in solution['metrics']:
                raise ValueError(f"Métrica ausente: {metric}")
        

    def export_results(self, results):
        """Exporta resultados para Excel"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Resultados"
        
        # Cabeçalhos ajustados para maior clareza
        headers = [
            "Ordem", "Padrão", "Layout", "Camadas", 
            "Comprimento Total (m)", "Aproveitamento (%)",
            "P", "M", "G", "GG",
            "Custo Total (R$)", 
            "Tecido Total (m)",
            "Desperdício (m²)",
            "Desperdício (m)",
            "Desperdício (%)"  
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
                        layout['num_layers'],
                        layout['length_meters'],
                        f"{layout['utilization']*100:.2f}",  
                        layout['production_per_size'].get('P', 0),
                        layout['production_per_size'].get('M', 0),
                        layout['production_per_size'].get('G', 0),
                        layout['production_per_size'].get('GG', 0),
                        f"R$ {result['metrics']['total_cost']:.2f}",  
                        f"{result['metrics']['fabric_meters']:.3f}",
                        f"{result['metrics']['fabric_waste_area']:.3f}",
                        f"{result['metrics']['fabric_waste_meters']:.3f}",
                        f"{result['metrics']['total_waste_percentage']:.2f}%"  
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

    

    def export_results_json(self, results):
        """Exporta resultados para arquivo JSON"""
        try:
            output = []
            for order_id, result in results.items():
                if result:
                    # Função auxiliar para converter Variable para int
                    def convert_variable(value):
                        if hasattr(value, 'solution_value'):
                            return int(value.solution_value())
                        return int(value)

                    # Converte e formata as métricas
                    metrics = {
                        "fabric_meters": round(float(result['metrics']['fabric_meters']), 3),
                        "fabric_cost": round(float(result['metrics']['fabric_cost']), 2),
                        "cutting_cost": round(float(result['metrics']['cutting_cost']), 2),
                        "layout_cost": round(float(result['metrics']['layout_cost']), 2),
                        "layer_cost": round(float(result['metrics']['layer_cost']), 2),
                        "waste_cost": round(float(result['metrics']['waste_cost']), 2),
                        "total_cost": round(sum([
                            float(result['metrics']['fabric_cost']),
                            float(result['metrics']['cutting_cost']),
                            float(result['metrics']['layout_cost']),
                            float(result['metrics']['layer_cost']),
                            float(result['metrics']['waste_cost'])
                        ]), 2),
                        "fabric_waste_area": round(float(result['metrics']['fabric_waste_area']), 4),
                        "fabric_waste_meters": round(float(result['metrics']['fabric_waste_meters']), 4),
                        "total_waste_percentage": round(float(result['metrics']['total_waste_percentage']), 2)
                    }

                    output_item = {
                        "order_id": order_id,
                        "pattern": result['pattern'],
                        "metrics": metrics,
                        "demand": {
                            size: convert_variable(qty) 
                            for size, qty in result.get('demand', {}).items()
                        },
                        "production": {
                            size: convert_variable(qty) 
                            for size, qty in result.get('production', {}).items()
                        },
                        "overproduction": {
                            size: convert_variable(qty) 
                            for size, qty in result.get('overproduction', {}).items()
                        },
                        "layouts_used": []
                    }

                    # Processa cada layout usado
                    for layout in result.get('layouts_used', []):
                        layout_info = {
                            "layout_id": layout['layout_id'],
                            "num_layers": convert_variable(layout['num_layers']),
                            "length_meters": round(float(layout['length_meters']), 3),
                            "utilization": round(float(layout['utilization']), 4),
                            "waste_area": round(float(layout['waste_area']), 4),
                            "production_per_size": {
                                size: convert_variable(qty) 
                                for size, qty in layout['production_per_size'].items()
                            },
                            "costs": {
                                "fabric_cost": round(float(layout['costs']['fabric_cost']), 2),
                                "cutting_cost": round(float(layout['costs']['cutting_cost']), 2),
                                "layout_cost": round(float(layout['costs']['layout_cost']), 2)
                            }
                        }
                        output_item["layouts_used"].append(layout_info)

                    output.append(output_item)

                # Escreve o resultado em um arquivo JSON
                with open('results.json', 'w', encoding='utf-8') as f:
                    json.dump(output, f, indent=4, ensure_ascii=False)

        except Exception as e:
            logging.error(f"Erro ao exportar resultados JSON: {str(e)}")
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