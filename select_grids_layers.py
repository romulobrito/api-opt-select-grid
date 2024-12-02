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
    """Utility class for unit conversions and metric calculations in the layout optimization process.

    This class provides static methods to handle various unit conversions commonly used
    in fabric cutting layouts, such as converting between millimeters and meters,
    or square centimeters to square meters. It also includes methods to calculate
    layout metrics and waste metrics with proper unit conversions.
    """

    @staticmethod
    def mm_to_m(value):
        """Convert millimeters to meters.

        Args:
            value (float): Length in millimeters

        Returns:
            float: Length in meters
        """
        return value / 1000.0

    @staticmethod
    def cm2_to_m2(value):
        """Convert square centimeters to square meters.

        Args:
            value (float): Area in square centimeters

        Returns:
            float: Area in square meters
        """
        return value / 10000.0

    @staticmethod
    def calculate_layout_metrics(layout, num_layers):
        """Calculate all layout metrics with correct units.

        Converts all measurements to standard units (meters, square meters)
        and calculates derived metrics for the layout.

        Args:
            layout (dict): Layout data containing measurements in original units
            num_layers (int): Number of fabric layers in the layout

        Returns:
            dict: Dictionary containing all converted metrics:
                - length_meters (float): Layout length in meters
                - waste_area_m2 (float): Waste area in square meters
                - total_area_m2 (float): Total area in square meters
                - perimeter_meters (float): Total perimeter in meters
                - fabric_width_m (float): Fabric width in meters
                - utilization (float): Layout utilization percentage
                - num_layers (int): Number of layers
        """
        return {
            "length_meters": UnitConverter.mm_to_m(layout["layout_length"]),
            "waste_area_m2": UnitConverter.cm2_to_m2(layout["waste_area"]),
            "total_area_m2": UnitConverter.cm2_to_m2(layout["total_area"]),
            "perimeter_meters": UnitConverter.mm_to_m(layout["total_perimeter"]),
            "fabric_width_m": UnitConverter.mm_to_m(layout["fabric_width"]),
            "utilization": layout["utilization"],
            "num_layers": num_layers,
        }

    @staticmethod
    def calculate_waste_metrics(layout):
        """Calculate waste-related metrics for a layout.

        Converts waste measurements to standard units and calculates
        derived waste metrics such as linear waste and waste percentage.

        Args:
            layout (dict): Layout data containing:
                - total_area (float): Total area in mm²
                - waste_area (float): Waste area in mm²
                - fabric_width (float): Fabric width in mm

        Returns:
            dict: Dictionary containing waste metrics:
                - fabric_waste_area (float): Waste area in m²
                - fabric_waste_meters (float): Linear waste in meters
                - total_waste_percentage (float): Waste percentage

        Raises:
            Exception: If any calculation errors occur
        """
        try:
            # Convert areas from mm² to m²
            total_area_m2 = layout["total_area"] / 10000  # mm² to m²
            waste_area_m2 = layout["waste_area"] / 10000  # mm² to m²
            fabric_width_m = layout["fabric_width"] / 1000  # mm to m

            # Calculate linear waste in meters
            waste_meters = waste_area_m2 / fabric_width_m if fabric_width_m > 0 else 0

            # Calculate waste percentage
            waste_percentage = (
                (layout["waste_area"] / layout["total_area"]) * 100
                if layout["total_area"] > 0
                else 0
            )

            return {
                "fabric_waste_area": waste_area_m2,
                "fabric_waste_meters": waste_meters,
                "total_waste_percentage": waste_percentage,
            }

        except Exception as e:
            logging.error(f"Error calculating waste metrics: {str(e)}")
            raise


class LayoutOptimizer:
    """Layout optimization engine for fabric cutting problems.

    This class handles the optimization of fabric cutting layouts, considering multiple
    constraints such as demand fulfillment, fabric waste minimization, and production costs.

    Attributes:
        config (dict): General configuration parameters
        layouts (list): Available cutting layouts
        fabrics (dict): Fabric specifications indexed by fabric name
        pieces (list): Pieces to be produced with their demands
        sizes (list): All available sizes sorted
        overproduction_penalty (float): Penalty factor for overproduction
        unit_waste_cost (float): Cost per unit of waste
        optimality_gap (float): Solver optimality gap
        max_memory_mb (int): Maximum memory usage for solver
        layout_change_penalty (float): Penalty for changing layouts
        min_layers_per_layout (int): Minimum number of layers per layout
        waste_penalty_factor (float): Factor to penalize waste in objective function
    """

    def __init__(self, input_data):
        """Initialize the layout optimizer with input data.

        Args:
            input_data (dict): Input data containing:
                - general_configuration (dict): Optimization parameters
                - layouts (list): Available cutting layouts
                - fabrics (list): Fabric specifications
                - pieces (list): Pieces to be produced

        Raises:
            ValueError: If required data is missing or invalid
        """
        # Load configuration and data
        self.config = input_data["general_configuration"]
        self.layouts = input_data.get("layout", input_data.get("layouts", []))
        self.fabrics = {f["fabric"]: f for f in input_data["fabrics"]}
        self.pieces = input_data["pieces"]

        # Extract sizes from all possible sources
        sizes_from_pieces = set()
        sizes_from_layouts = set()

        # Extract sizes from pieces (demand)
        for piece in self.pieces:
            if "quantity" in piece:
                sizes_from_pieces.update(piece["quantity"].keys())

        # Extract sizes from layouts
        for layout in self.layouts:
            for piece in layout["pieces"]:
                if "size_grade" in piece:
                    sizes_from_layouts.update(piece["size_grade"].keys())

        # Combine all found sizes
        self.sizes = sorted(sizes_from_pieces.union(sizes_from_layouts))

        if not self.sizes:
            raise ValueError("Could not extract sizes from input data")

        logging.info(f"Automatically extracted sizes: {self.sizes}")

        # Initialize optimization parameters
        self.overproduction_penalty = self.config.get("overproduction_percentage", 0.05)
        self.unit_waste_cost = self.config.get("waste_cost", 0.1)
        self.optimality_gap = self.config.get("optimality_gap", 0.01)
        self.max_memory_mb = self.config.get("max_memory_mb", 2048)

        # New parameters
        self.layout_change_penalty = self.config.get("layout_change_penalty", 1000)
        self.min_layers_per_layout = self.config.get("min_layers_per_layout", 5)
        self.waste_penalty_factor = self.config.get("waste_penalty_factor", 2.0)

        # Log loaded parameters
        logging.info("Optimization parameters:")
        logging.info(f"  Layout change penalty: {self.layout_change_penalty}")
        logging.info(f"  Min layers per layout: {self.min_layers_per_layout}")
        logging.info(f"  Waste penalty factor: {self.waste_penalty_factor}")

        # Validate parameters
        if self.layout_change_penalty < 0:
            raise ValueError("layout_change_penalty must be non-negative")

        if self.min_layers_per_layout < 1:
            raise ValueError("min_layers_per_layout must be at least 1")

        if self.waste_penalty_factor < 0:
            raise ValueError("waste_penalty_factor must be non-negative")

        # Calculate price per square meter for each fabric
        for fabric in self.fabrics.values():
            fabric["price_per_square_meter"] = fabric["price_per_linear_meter"] / (
                fabric["fabric_width"] / 1000
            )
            logging.info(
                f"Price per m² for fabric {fabric['fabric']}: ${fabric['price_per_square_meter']:.2f}"
            )

    def optimize_production(self):
        """Optimize production for all pieces in the order.

        This method processes each piece in the order separately, creating an optimization
        problem for each one and collecting the results. It ensures all metrics are properly
        calculated and validated.

        Returns:
            dict: Results for each order, containing:
                - metrics (dict): Production metrics including costs and waste
                - layouts_used (list): Selected layouts and their configurations
                - production (dict): Pieces produced by size
                - overproduction (dict): Excess production by size

        Example result structure:
            {
                'order_1': {
                    'metrics': {
                        'fabric_meters': 100.5,
                        'fabric_cost': 1000.0,
                        ...
                    },
                    'layouts_used': [...],
                    'production': {'S': 50, 'M': 60, ...},
                    'overproduction': {'S': 2, 'M': 1, ...}
                },
                ...
            }
        """
        results = {}

        for i, piece in enumerate(self.pieces, 1):
            order_id = f"order_{i}"
            pattern = piece["pattern"]
            logging.info(f"Processing {pattern} (Order {order_id})")

            # Create order structure with required parameters
            order = {
                "id": order_id,
                "pattern": pattern,
                "pieces": [piece],
                "fabric_width": self.fabrics[piece["fabrics"][0]]["fabric_width"],
                "max_layers": self.fabrics[piece["fabrics"][0]]["max_layers"],
                "max_length": self.config["max_total_length"],
            }

            # Optimize individual order
            result = self.optimize_order(order)
            if result:
                # Initialize default metrics to ensure all fields exist
                metrics = {
                    "fabric_meters": 0.0,  # Total fabric length in meters
                    "fabric_cost": 0.0,  # Cost of fabric used
                    "cutting_cost": 0.0,  # Cost of cutting operations
                    "layout_cost": 0.0,  # Cost of layout setup
                    "layer_cost": 0.0,  # Cost related to number of layers
                    "waste_cost": 0.0,  # Cost of wasted material
                    "total_cost": 0.0,  # Sum of all costs
                    "fabric_waste_area": 0.0,  # Waste area in square meters
                    "fabric_waste_meters": 0.0,  # Waste in linear meters
                    "total_waste_percentage": 0.0,  # Percentage of total waste
                }

                # Update with calculated values from optimization
                metrics.update(result["metrics"])

                # Recalculate total cost to ensure consistency
                metrics["total_cost"] = sum(
                    [
                        metrics["fabric_cost"],
                        metrics["cutting_cost"],
                        metrics["layout_cost"],
                        metrics["layer_cost"],
                        metrics["waste_cost"],
                    ]
                )

                result["metrics"] = metrics
                results[order_id] = result

                # Final validation of metrics
                self._validate_metrics(result["metrics"])

        return results

    # TODO: Implementar validação de métricas de layout
    def _validate_layout_metrics(self, layout_data):
        """Validate the metrics of the layout.

        This method checks the validity of the layout metrics to ensure that
        the length in meters is greater than zero and consistent with the
        length provided in millimeters. It raises a ValueError if any
        validation fails.

        Args:
            layout_data (dict): Layout metrics containing:
                - length_meters (float): Length of the layout in meters
                - layout_length (float): Length of the layout in millimeters

        Raises:
            ValueError: If the length is invalid or inconsistent.
        """
        # Length in meters must be greater than 0
        if layout_data["length_meters"] <= 0:
            raise ValueError(f"Invalid length: {layout_data['length_meters']}")

        # Length in meters must be consistent with the value in millimeters
        expected_length = layout_data["layout_length"] / 1000  # Convert mm to m
        if abs(layout_data["length_meters"] - expected_length) > 0.001:
            raise ValueError(
                f"Inconsistency in length: "
                f"expected {expected_length}m, "
                f"obtained {layout_data['length_meters']}m"
            )

    # TODO: Implementar validação de métricas de desperdício
    def _validate_waste_metrics(self, layout_data):
        """Validate waste metrics of the layout.

        This method checks the waste metrics calculated for a layout to ensure
        that the utilization is consistent with the calculated waste area. It
        raises a ValueError if any validation fails.

        Args:
            layout_data (dict): Layout data containing:
                - id (str): Identifier for the layout
                - utilization (float): Reported utilization percentage
                - total_area_m2 (float): Total area of the layout in square meters

        Returns:
            dict: Calculated waste metrics including:
                - fabric_waste_area (float): Area of fabric wasted in square meters
                - fabric_waste_meters (float): Linear waste in meters

        Raises:
            ValueError: If the calculated utilization is inconsistent with the reported utilization.
        """
        try:
            # Convert units and calculate metrics
            metrics = UnitConverter.calculate_layout_metrics(layout_data, 1)
            waste_metrics = UnitConverter.calculate_waste_metrics(metrics)

            # Validate utilization
            calculated_utilization = 1 - (
                waste_metrics["fabric_waste_area"] / metrics["total_area_m2"]
            )
            if abs(calculated_utilization - layout_data["utilization"]) > 0.01:
                raise ValueError(
                    f"Inconsistency in utilization calculation for layout {layout_data['id']}: "
                    f"calculated {calculated_utilization:.4f}, reported {layout_data['utilization']:.4f}"
                )

            return waste_metrics

        except Exception as e:
            logging.error(f"Error validating waste metrics: {str(e)}")
            raise

    def _validate_metrics(self, metrics):
        """Validate the consistency of calculated metrics.

        This method checks various metrics to ensure they are valid and consistent.
        It raises a ValueError if any validation fails, including checks for negative
        values, total cost consistency, waste metrics, and more.

        Args:
            metrics (dict): Calculated metrics containing:
                - fabric_cost (float): Cost of fabric used
                - cutting_cost (float): Cost of cutting operations
                - layout_cost (float): Cost of layout setup
                - layer_cost (float): Cost related to the number of layers
                - waste_cost (float): Cost of wasted material
                - total_cost (float): Total cost of production
                - fabric_waste_area (float): Area of fabric wasted in square meters
                - fabric_waste_meters (float): Linear waste in meters
                - total_waste_percentage (float): Percentage of total waste
                - fabric_meters (float): Total fabric length in meters

        Raises:
            ValueError: If any metric is invalid or inconsistent.
        """
        try:
            # Check for negative values
            for key, value in metrics.items():
                if value < 0:
                    raise ValueError(f"Metric {key} has a negative value: {value}")

            # Check consistency of total cost
            expected_total = sum(
                [
                    metrics["fabric_cost"],
                    metrics["cutting_cost"],
                    metrics["layout_cost"],
                    metrics["layer_cost"],
                    metrics["waste_cost"],
                ]
            )

            if abs(expected_total - metrics["total_cost"]) > 0.01:
                raise ValueError(
                    f"Inconsistency in total cost: "
                    f"expected R${expected_total:.2f}, "
                    f"obtained R${metrics['total_cost']:.2f}"
                )

            # Check consistency of waste metrics
            if metrics["fabric_waste_meters"] > metrics["fabric_meters"]:
                raise ValueError("Waste meters exceed total fabric meters")

            # Specific validation for waste metrics
            if metrics["fabric_waste_area"] > 0:
                expected_waste_meters = metrics["fabric_waste_area"] / (
                    self.fabrics[self.pieces[0]["fabrics"][0]]["fabric_width"] / 1000
                )
                if abs(metrics["fabric_waste_meters"] - expected_waste_meters) > 0.0001:
                    raise ValueError(
                        f"Inconsistency in waste meters: "
                        f"expected {expected_waste_meters:.6f}m, "
                        f"obtained {metrics['fabric_waste_meters']:.6f}m"
                    )

            if metrics["fabric_meters"] <= 0:
                raise ValueError("Fabric meters must be greater than zero")

            if metrics["total_cost"] <= 0:
                raise ValueError("Total cost must be greater than zero")

            # Validations for waste metrics
            if metrics["fabric_waste_area"] < 0:
                raise ValueError("Waste area cannot be negative")

            if metrics["fabric_waste_meters"] < 0:
                raise ValueError("Waste meters cannot be negative")

            if not (0 <= metrics["total_waste_percentage"] <= 100):
                raise ValueError(
                    f"Waste percentage must be between 0 and 100, found: "
                    f"{metrics['total_waste_percentage']:.2f}%"
                )

        except Exception as e:
            logging.error(f"Error validating metrics: {str(e)}")
            raise

    def _validate_costs(self, layout_data):
        """Validate the costs associated with the layout.

        This method checks the calculated costs of a layout against expected values
        based on the fabric specifications and layout dimensions. It raises a ValueError
        if any cost is inconsistent with the expected value.

        Args:
            layout_data (dict): Layout data containing:
                - fabric (str): Name of the fabric used
                - num_layers (int): Number of layers in the layout
                - layout_length (float): Length of the layout in millimeters
                - waste_area (float): Area of waste in mm²
                - fabric_width (float): Width of the fabric in mm
                - total_perimeter (float): Total perimeter of the layout in mm
                - costs (dict): Calculated costs including:
                    - fabric_cost (float): Cost of fabric used
                    - cutting_cost (float): Cost of cutting operations
                    - layout_cost (float): Cost of layout setup
                    - layer_cost (float): Cost related to the number of layers
                    - waste_cost (float): Cost of wasted material

        Raises:
            ValueError: If any calculated cost is inconsistent with the expected cost.
        """
        try:
            fabric = self.fabrics[layout_data["fabric"]]
            tolerance = 0.01
            num_layers = layout_data["num_layers"]
            length_meters = layout_data["layout_length"] / 1000  # Convert mm to m

            # Conversions for waste calculation
            waste_area_m2 = layout_data["waste_area"] / 1_000_000  # Convert mm² to m²
            fabric_width_m = layout_data["fabric_width"] / 1000  # Convert mm to m
            waste_meters = waste_area_m2 / fabric_width_m

            # Expected costs
            expected_costs = {
                "fabric_cost": length_meters
                * fabric["price_per_linear_meter"]
                * num_layers,
                "cutting_cost": (layout_data["total_perimeter"] / 1000)
                * fabric["cost_per_cut_meter"]
                * num_layers,
                "layout_cost": length_meters
                * fabric["cost_per_layout_meter"]
                * num_layers,
                "layer_cost": fabric["cost_per_layer"] * num_layers,
                "waste_cost": waste_meters
                * fabric["price_per_linear_meter"]
                * num_layers,
            }

            # Validation
            for cost_type, expected in expected_costs.items():
                actual = layout_data["costs"].get(cost_type, 0)
                if abs(actual - expected) > tolerance:
                    raise ValueError(
                        f"Error in calculating {cost_type}: "
                        f"expected R${expected:.2f}, "
                        f"obtained R${actual:.2f}"
                    )

        except Exception as e:
            logging.error(f"Error validating costs: {str(e)}")
            raise

    def _calculate_waste_cost(self, layout, num_layers):
        """Calculate the waste cost associated with a layout.

        This method computes the cost of waste based on the layout's waste area,
        fabric specifications, and the number of layers. It raises a ValueError
        if any validation checks fail, such as negative costs or excessive waste.

        Args:
            layout (dict): Layout data containing:
                - fabric (str): Name of the fabric used
                - waste_area (float): Area of waste in mm²
                - fabric_width (float): Width of the fabric in mm
                - layout_length (float): Length of the layout in mm
            num_layers (int): Number of layers in the layout

        Returns:
            dict: Dictionary containing:
                - waste_cost (float): Calculated cost of waste
                - fabric_waste_area (float): Area of fabric wasted in m²
                - fabric_waste_meters (float): Linear waste in meters

        Raises:
            ValueError: If waste cost is negative or waste meters exceed layout length.
        """
        try:
            fabric = self.fabrics[layout["fabric"]]

            # Correct unit conversions
            waste_area_m2 = UnitConverter.cm2_to_m2(layout["waste_area"] / 100)  # mm² -> cm² -> m²
            fabric_width_m = UnitConverter.mm_to_m(layout["fabric_width"])  # mm -> m
            waste_meters = waste_area_m2 / fabric_width_m

            # Calculate waste cost
            waste_cost = waste_meters * fabric["price_per_linear_meter"] * num_layers

            # Validations
            if waste_cost < 0:
                raise ValueError(f"Negative waste cost: {waste_cost}")
            if waste_meters > layout["layout_length"] / 1000:
                raise ValueError(
                    f"Waste meters ({waste_meters}) exceed layout length ({layout['layout_length'] / 1000})"
                )

            return {
                "waste_cost": waste_cost,
                "fabric_waste_area": waste_area_m2,
                "fabric_waste_meters": waste_meters,
            }
        except Exception as e:
            logging.error(f"Error calculating waste cost: {str(e)}")
            raise

    # TODO: Implementar validação de métricas de desperdício
    def _calculate_layout_metrics(self, layout_data):
        """Calculate metrics for a specific layout.

        This method computes various cost metrics associated with a layout,
        including fabric cost, cutting cost, layout cost, layer cost, and waste
        metrics. It raises a ValueError if any calculation fails.

        Args:
            layout_data (dict): Layout data containing:
                - fabric (str): Name of the fabric used
                - num_layers (int): Number of layers in the layout
                - length_meters (float): Length of the layout in mm

        Returns:
            dict: Dictionary containing:
                - fabric_cost (float): Cost of fabric used
                - cutting_cost (float): Cost of cutting operations
                - layout_cost (float): Cost of layout setup
                - layer_cost (float): Cost related to the number of layers
                - waste_cost (float): Cost of waste
                - fabric_waste_area (float): Area of fabric wasted in m²
                - fabric_waste_meters (float): Linear waste in meters

        Raises:
            ValueError: If any calculation errors occur.
        """
        try:
            fabric = self.fabrics[layout_data["fabric"]]
            num_layers = layout_data["num_layers"]

            # Basic calculations
            length_meters = UnitConverter.mm_to_m(layout_data["length_meters"])  # ao invés de / 1000
            fabric_cost = length_meters * fabric["price_per_linear_meter"] * num_layers
            cutting_cost = length_meters * fabric["cost_per_cut_meter"] * num_layers
            layout_cost = length_meters * fabric["cost_per_layout_meter"]
            layer_cost = fabric["cost_per_layer"] * num_layers

            # Calculate waste metrics
            waste_metrics = self._calculate_waste_cost(layout_data, num_layers)

            return {
                "fabric_cost": fabric_cost,
                "cutting_cost": cutting_cost,
                "layout_cost": layout_cost,
                "layer_cost": layer_cost,
                "waste_cost": waste_metrics["waste_cost"],
                "fabric_waste_area": waste_metrics["fabric_waste_area"],
                "fabric_waste_meters": waste_metrics["fabric_waste_meters"],
            }
        except Exception as e:
            logging.error(f"Error calculating layout metrics: {str(e)}")
            raise

    def preprocess_layouts(self, demand, fabric_width):
        """Preprocess and filter layouts compatible with the demand and fabric width.

        This method filters the available layouts based on the specified fabric width
        and checks if they contain the necessary patterns for the given demand. It also
        calculates the efficiency of each layout per size and returns the filtered layouts.

        Args:
            demand (dict): Demand data containing:
                - pieces (list): List of pieces required, each with a pattern and size grade.
            fabric_width (float): Width of the fabric to filter layouts.

        Returns:
            list: A list of filtered layouts that match the demand and fabric width,
                each with calculated efficiency metrics.
        """
        filtered_layouts = []
        for layout in self.layouts:
            # Filter by fabric width
            if layout["fabric_width"] != fabric_width:
                continue

            # Check if the layout contains the necessary patterns
            layout_patterns = [p["pattern"] for p in layout["pieces"]]
            order_patterns = [p["pattern"] for p in demand["pieces"]]
            if not set(order_patterns).issubset(set(layout_patterns)):
                continue

            # Calculate efficiency per size
            layout["efficiency"] = {}
            for size in self.sizes:
                total_pieces_size = sum(
                    p["size_grade"].get(size, 0) for p in layout["pieces"]
                )
                pieces_per_meter = total_pieces_size / (layout["layout_length"] / 1000)
                efficiency = pieces_per_meter * layout["utilization"]
                layout["efficiency"][size] = efficiency

            filtered_layouts.append(layout)

        return filtered_layouts

    def _calculate_layout_costs(self, layout, num_layers):
        """Calculate the costs associated with a layout.

        This method computes various costs related to a layout, including fabric cost,
        cutting cost, layout cost, layer cost, and waste cost. It raises a ValueError
        if any calculation fails.

        Args:
            layout (dict): Layout data containing:
                - fabric (str): Name of the fabric used
                - total_area (float): Total area of the layout in mm²
                - waste_area (float): Area of waste in mm²
                - fabric_width (float): Width of the fabric in mm
            num_layers (int): Number of layers in the layout

        Returns:
            dict: Dictionary containing:
                - fabric_meters (float): Total fabric length in meters
                - fabric_cost (float): Cost of fabric used
                - cutting_cost (float): Cost of cutting operations
                - layout_cost (float): Cost of layout setup
                - layer_cost (float): Cost related to the number of layers
                - waste_cost (float): Cost of waste
                - fabric_waste_area (float): Area of fabric wasted in m²
                - fabric_waste_meters (float): Linear waste in meters
                - total_waste_percentage (float): Percentage of total waste

        Raises:
            ValueError: If any calculation errors occur.
        """
        try:
            fabric = self.fabrics[layout["fabric"]]

            # First, calculate basic metrics of the layout
            base_metrics = UnitConverter.calculate_layout_metrics(layout, num_layers)

            # Prepare data for waste calculation
            waste_data = {
                "total_area": layout["total_area"],  # Original area in mm²
                "waste_area": layout["waste_area"],  # Waste area in mm²
                "fabric_width": layout["fabric_width"],  # Width in mm
            }

            # Calculate waste metrics
            waste_metrics = UnitConverter.calculate_waste_metrics(waste_data)

            # Calculate costs
            costs = {
                "fabric_meters": base_metrics["length_meters"] * num_layers,
                "fabric_cost": base_metrics["length_meters"]
                * fabric["price_per_linear_meter"]
                * num_layers,
                "cutting_cost": base_metrics["perimeter_meters"]
                * fabric["cost_per_cut_meter"]
                * num_layers,
                "layout_cost": base_metrics["length_meters"]
                * fabric["cost_per_layout_meter"]
                * num_layers,
                "layer_cost": fabric["cost_per_layer"] * num_layers,
                "waste_cost": waste_metrics["fabric_waste_meters"]
                * fabric["price_per_linear_meter"]
                * num_layers,
            }

            # Add waste metrics to costs
            costs.update(
                {
                    "fabric_waste_area": waste_metrics["fabric_waste_area"]
                    * num_layers,
                    "fabric_waste_meters": waste_metrics["fabric_waste_meters"]
                    * num_layers,
                    "total_waste_percentage": waste_metrics["total_waste_percentage"],
                }
            )

            return costs

        except Exception as e:
            logging.error(
                f"Error calculating costs for layout {layout.get('id')}: {str(e)}"
            )
            raise

    def _process_solution(
        self, solver, x, y, overproduction, filtered_layouts, demand_quantity, pattern
    ):
        """Process the solver's solution and calculate all metrics.

        This method extracts the solution values from the solver, calculates various
        metrics related to the production layout, and prepares the results for output.

        Args:
            solver: The solver instance used to find the solution.
            x (dict): Decision variable values indicating the number of layers for each layout.
            y (dict): Decision variable values indicating whether a layout is used.
            overproduction (dict): Overproduction values for each size.
            filtered_layouts (list): List of layouts that were filtered for compatibility.
            demand_quantity (dict): Quantity of demand for each size.
            pattern (str): The pattern being processed.

        Returns:
            dict: A dictionary containing:
                - pattern (str): The pattern being processed.
                - metrics (dict): Calculated metrics including costs and waste.
                - demand (dict): Demand quantities for each size.
                - production (dict): Production quantities for each size.
                - overproduction (dict): Overproduction quantities for each size.
                - layouts_used (list): Information about the layouts used.

        Raises:
            ValueError: If any processing errors occur.
        """
        try:
            # Capture the values of the variables immediately after the solution
            layout_solutions = {
                layout["id"]: x[layout["id"]].solution_value()
                for layout in filtered_layouts
            }

            result = {
                "pattern": pattern,
                "metrics": {
                    "fabric_meters": 0.0,
                    "fabric_cost": 0.0,
                    "cutting_cost": 0.0,
                    "layout_cost": 0.0,
                    "layer_cost": 0.0,
                    "waste_cost": 0.0,
                    "total_cost": 0.0,
                    "fabric_waste_area": 0.0,
                    "fabric_waste_meters": 0.0,
                    "total_waste_percentage": 0.0,
                },
                "demand": {size: demand_quantity.get(size, 0) for size in self.sizes},
                "production": {size: 0 for size in self.sizes},
                "overproduction": {size: 0 for size in self.sizes},
                "layouts_used": [],
            }

            total_area = 0
            total_waste_area = 0

            # Process each used layout
            for layout in filtered_layouts:
                num_layers = int(layout_solutions[layout["id"]])
                if num_layers > 0:
                    # Calculate basic metrics of the layout
                    metrics = UnitConverter.calculate_layout_metrics(layout, num_layers)

                    # Calculate waste metrics
                    waste_metrics = UnitConverter.calculate_waste_metrics(
                        {
                            "total_area": layout["total_area"],
                            "waste_area": layout["waste_area"],
                            "fabric_width": layout["fabric_width"],
                        }
                    )

                    # Calculate layout costs
                    layout_costs = self._calculate_layout_costs(layout, num_layers)

                    # Prepare layout information
                    layout_info = {
                        "layout_id": layout["id"],
                        "num_layers": num_layers,
                        "length_meters": metrics["length_meters"],
                        "utilization": layout["utilization"],
                        "waste_area": waste_metrics["fabric_waste_area"],
                        "production_per_size": {size: 0 for size in self.sizes},
                        "costs": layout_costs,
                    }

                    # Calculate production by size
                    for size in self.sizes:
                        if size in layout["pieces"][0]["size_grade"]:
                            qty = layout["pieces"][0]["size_grade"][size] * num_layers
                            layout_info["production_per_size"][size] = qty
                            result["production"][size] += qty

                    # Update total metrics
                    result["metrics"]["fabric_meters"] += (
                        metrics["length_meters"] * num_layers
                    )
                    result["metrics"]["fabric_cost"] += layout_costs["fabric_cost"]
                    result["metrics"]["cutting_cost"] += layout_costs["cutting_cost"]
                    result["metrics"]["layout_cost"] += layout_costs["layout_cost"]
                    result["metrics"]["layer_cost"] += layout_costs["layer_cost"]
                    result["metrics"]["waste_cost"] += layout_costs["waste_cost"]
                    result["metrics"]["fabric_waste_area"] += (
                        waste_metrics["fabric_waste_area"] * num_layers
                    )
                    result["metrics"]["fabric_waste_meters"] += (
                        waste_metrics["fabric_waste_meters"] * num_layers
                    )

                    # Accumulate areas for total waste calculation
                    total_area += layout["total_area"] * num_layers
                    total_waste_area += layout["waste_area"] * num_layers

                    result["layouts_used"].append(layout_info)

            # Calculate total waste percentage
            if total_area > 0:
                result["metrics"]["total_waste_percentage"] = (
                    total_waste_area / total_area
                ) * 100

            # Calculate total cost
            result["metrics"]["total_cost"] = sum(
                [
                    result["metrics"]["fabric_cost"],
                    result["metrics"]["cutting_cost"],
                    result["metrics"]["layout_cost"],
                    result["metrics"]["layer_cost"],
                    result["metrics"]["waste_cost"],
                ]
            )

            # Calculate final overproduction
            for size in self.sizes:
                if size in demand_quantity:
                    result["overproduction"][size] = max(
                        0, result["production"][size] - demand_quantity[size]
                    )

            return result

        except Exception as e:
            logging.error(f"Error processing the solution: {str(e)}")
            logging.error(traceback.format_exc())
            raise

    # TODO: Implementar validação de métricas de desperdício
    def _validate_final_results(self, result):
        """Final validation of the results.

        This method checks the final production results against the demand and validates
        the total costs to ensure consistency. It raises a ValueError if any validation
        checks fail.

        Args:
            result (dict): The results dictionary containing:
                - production (dict): Production quantities for each size.
                - demand (dict): Demand quantities for each size.
                - metrics (dict): Calculated metrics including costs.

        Raises:
            ValueError: If production is insufficient or excessive, or if total costs are inconsistent.
        """
        try:
            # Validate production vs demand
            for size in self.sizes:
                if result["production"][size] < result["demand"][size]:
                    raise ValueError(f"Insufficient production for size {size}")
                if result["production"][size] > result["demand"][size] * 1.05:
                    raise ValueError(f"Excessive overproduction for size {size}")

            # Validate total costs
            total_cost = sum(
                [
                    result["metrics"]["fabric_cost"],
                    result["metrics"]["cutting_cost"],
                    result["metrics"]["layout_cost"],
                    result["metrics"]["layer_cost"],
                    result["metrics"]["waste_cost"],
                ]
            )

            if abs(total_cost - result["metrics"]["total_cost"]) > 0.01:
                raise ValueError(
                    f"Inconsistency in total cost: "
                    f"calculated {total_cost:.2f}, "
                    f"reported {result['metrics']['total_cost']:.2f}"
                )

        except Exception as e:
            logging.error(f"Error in final validation: {str(e)}")
            raise

    def _validate_units(self, result):
        """Validate the units of the calculated metrics.

        This method checks the calculated metrics to ensure that they are consistent
        with the expected units. It raises a ValueError if any validation checks fail.

        Args:
            result (dict): The results dictionary containing:
                - metrics (dict): Calculated metrics including fabric meters.
                - layouts_used (list): Information about the layouts used.

        Raises:
            ValueError: If any unit validations fail.
        """
        try:
            # Existing validations
            if result["metrics"]["fabric_meters"] < 0:
                raise ValueError("Fabric meters cannot be negative")

            # New validations
            for layout in result["layouts_used"]:
                # Validate conversions
                length_m = UnitConverter.mm_to_m(layout["layout_length"])
                waste_m2 = UnitConverter.cm2_to_m2(layout["waste_area"])

                # Compare with calculated values
                if abs(layout["length_meters"] - length_m) > 0.001:
                    raise ValueError(
                        f"Inconsistency in length conversion for layout {layout['layout_id']}"
                    )

                if abs(layout["waste_area"] - waste_m2) > 0.001:
                    raise ValueError(
                        f"Inconsistency in area conversion for layout {layout['layout_id']}"
                    )

        except Exception as e:
            logging.error(f"Error in unit validation: {str(e)}")
            raise

    def optimize_order(self, order):
        """
        Optimize a specific order allowing multiple layouts.

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

        Args:
            order (dict): The order data containing:
                - pieces (list): List of pieces to be produced.
                - max_layers (int): Maximum number of layers allowed for the order.
                - pattern (str): The pattern associated with the order.

        Returns:
            dict: A dictionary containing the optimization results, including metrics and production details.

        Raises:
            ValueError: If any validation or optimization errors occur.
        """
        try:
            logging.info("=== Starting Optimization ===")
            logging.info(f"Demand by size: {order['pieces'][0]['quantity']}")
            self._validate_input_data(order)

            piece = order["pieces"][0]
            pattern = order["pattern"]
            demand_quantity = piece["quantity"]
            fabric = piece["fabrics"][0]

            logging.info(f"Starting optimization for {pattern}")
            logging.info(f"Demand: {demand_quantity}")

            # Filter and sort layouts
            filtered_layouts = [
                layout
                for layout in self.layouts
                if layout["fabric"] == fabric
                and any(p["pattern"] == pattern for p in layout["pieces"])
            ]
            filtered_layouts.sort(key=lambda x: x["utilization"], reverse=True)

            # Validate sizes in all filtered layouts
            for layout in filtered_layouts:
                self._validate_layout_sizes(layout)

            # Log available layouts
            for layout in filtered_layouts:
                logging.info(
                    f"Layout {layout['id']}: {layout['utilization']*100:.1f}% utilization"
                )
                logging.info("  Pieces per layer:")
                for size, qty in layout["pieces"][0]["size_grade"].items():
                    logging.info(f"    {size}: {qty}")

            # Solver setup
            solver = pywraplp.Solver.CreateSolver("SCIP")
            solver.SetTimeLimit(
                self.config.get("solver_time_limit", 1) * 60 * 1000
            )  # Convert minutes to milliseconds
            scip_params = (
                f"limits/gap = {self.optimality_gap}\n"
                f"limits/memory = {self.max_memory_mb}\n"
                "display/verblevel = 4\n"  # Detailed log level
                "timing/clocktype = 1\n"  # Use CPU time
            )
            solver.SetSolverSpecificParametersAsString(scip_params)

            # Decision variables
            x = {}  # Number of layers per layout
            y = {}  # Binary variable indicating if the layout is used

            # Initialize variables for each layout
            for layout in filtered_layouts:
                x[layout["id"]] = solver.IntVar(
                    0, order["max_layers"], f'x_{layout["id"]}'
                )
                y[layout["id"]] = solver.IntVar(0, 1, f'y_{layout["id"]}')

                # Relate x and y: if y=0, x must be 0
                solver.Add(x[layout["id"]] <= order["max_layers"] * y[layout["id"]])

                # Ensure at least one layout is used
                solver.Add(
                    x[layout["id"]] >= self.min_layers_per_layout * y[layout["id"]]
                )

                # If not used (y=0), must have 0 layers
                solver.Add(x[layout["id"]] <= order["max_layers"] * y[layout["id"]])

            # Create variables for overproduction
            overproduction = {}
            for size in self.sizes:
                overproduction[size] = solver.NumVar(
                    0, solver.infinity(), f"over_{size}"
                )

            # Demand constraints by size
            for size in self.sizes:
                if size in demand_quantity:
                    # Sum production from all layouts
                    total_production = solver.Sum(
                        [
                            x[layout["id"]]
                            * layout["pieces"][0]["size_grade"].get(size, 0)
                            for layout in filtered_layouts
                        ]
                    )

                    # Ensure minimum production
                    solver.Add(total_production >= demand_quantity[size])

                    # Limit excess production
                    solver.Add(
                        total_production
                        <= demand_quantity[size] * (1 + self.overproduction_penalty)
                    )  # 5% maximum

                    logging.info(f"Configured constraint for size {size}:")
                    logging.info(f"  Demand: {demand_quantity[size]}")
                    logging.info(f"  Minimum production: {demand_quantity[size]}")
                    logging.info(
                        f"  Maximum production: {demand_quantity[size] * 1.05}"
                    )

            # Corrected objective function
            objective = solver.Sum(
                [
                    x[layout["id"]]
                    * UnitConverter.cm2_to_m2(
                        float(layout["waste_area"])
                    )  # Convert to float and m²
                    + solver.Sum(
                        [
                            overproduction[size]
                            * 0.01
                            * float(layout["waste_area"])
                            / float(layout["total_area"])
                            for size in self.sizes
                        ]
                    )
                    for layout in filtered_layouts
                ]
            )

            solver.Minimize(objective)

            # Solve the problem
            status = solver.Solve()
            logging.info(f"Solution status: {status}")

            if status == solver.OPTIMAL:
                logging.info("Optimal solution found!")
                logging.info(f"Objective value: {solver.Objective().Value():.6f}")
                logging.info(f"Best bound: {solver.Objective().BestBound():.6f}")
                gap = abs(
                    solver.Objective().Value() - solver.Objective().BestBound()
                ) / abs(solver.Objective().Value())
                logging.info(f"Optimality gap: {gap*100:.6f}%")
            elif status == solver.FEASIBLE:
                logging.info("Feasible solution found (not necessarily optimal)")
                logging.info(f"Optimality gap: {gap*100:.6f}%")
            else:
                logging.warning("No feasible solution found")
                return None

            # Log the found solution
            if status in [solver.OPTIMAL, solver.FEASIBLE]:
                logging.info("\nSolution details:")
                total_production = {size: 0 for size in self.sizes}

                for layout in filtered_layouts:
                    num_layers = int(x[layout["id"]].solution_value())
                    if num_layers > 0:
                        logging.info(f"\nLayout {layout['id']} selected:")
                        logging.info(f"  Number of layers: {num_layers}")
                        logging.info(f"  Utilization: {layout['utilization']*100:.1f}%")
                        logging.info("  Production by size:")

                        for size in self.sizes:
                            if size in layout["pieces"][0]["size_grade"]:
                                prod = (
                                    num_layers * layout["pieces"][0]["size_grade"][size]
                                )
                                total_production[size] += prod
                                logging.info(f"    {size}: {prod}")

                logging.info("\nTotal production by size:")
                for size in self.sizes:
                    if size in demand_quantity:
                        logging.info(
                            f"  {size}: {total_production[size]} (Demand: {demand_quantity[size]})"
                        )

                if (
                    status == pywraplp.Solver.OPTIMAL
                    or status == pywraplp.Solver.FEASIBLE
                ):
                    # Process solution
                    result = self._process_solution(
                        solver=solver,
                        x=x,
                        y=y,
                        overproduction=overproduction,
                        filtered_layouts=filtered_layouts,
                        demand_quantity=demand_quantity,
                        pattern=pattern,
                    )

                    # Add solver information
                    result["solver_info"] = {
                        "status": (
                            "optimal"
                            if status == pywraplp.Solver.OPTIMAL
                            else "feasible"
                        ),
                        "gap": gap,
                        "objective_value": solver.Objective().Value(),
                        "best_bound": solver.Objective().BestBound(),
                        "iterations": solver.iterations(),
                        "wall_time": solver.WallTime() / 1000.0,
                    }
                    self._validate_solution(result, demand_quantity)

                    logging.info("=== Optimization Results ===")
                    logging.info(f"Status: {status}")
                    logging.info(f"Objective value: {solver.Objective().Value()}")
                    logging.info("Detailed costs:")
                    for metric, value in result["metrics"].items():
                        logging.info(f"  {metric}: {value:.2f}")

                    return result

            return None

        except Exception as e:
            logging.error(f"Error in optimization: {str(e)}")
            import traceback

            logging.error(traceback.format_exc())
            return None

    def _validate_input_data(self, order):
        """Validate input data for the order.

        This method checks the validity of the input order data, ensuring that it contains
        the necessary pieces and that the demand sizes are valid. It raises a ValueError
        if any validation checks fail.

        Args:
            order (dict): The order data containing:
                - pieces (list): List of pieces to be produced, each with a pattern, quantity, and fabrics.

        Raises:
            ValueError: If the order is missing pieces, if any piece data is incomplete,
                        or if the demand sizes are invalid.
        """
        if not order.get("pieces"):
            raise ValueError("Order must contain pieces")

        piece = order["pieces"][0]
        if not all(key in piece for key in ["pattern", "quantity", "fabrics"]):
            raise ValueError("Incomplete piece data")

        # Validate that all demand sizes are among the extracted sizes
        demand_sizes = set(piece["quantity"].keys())
        if not demand_sizes.issubset(set(self.sizes)):
            invalid_sizes = demand_sizes - set(self.sizes)
            raise ValueError(f"Invalid sizes in demand: {invalid_sizes}")

        # New validation: Check if layouts can produce all demanded sizes
        layouts_production_capacity = {}
        for layout in self.layouts:
            if layout["pieces"][0]["pattern"] == piece["pattern"]:
                for size, qty in layout["pieces"][0]["size_grade"].items():
                    if size in layouts_production_capacity:
                        layouts_production_capacity[size] = max(
                            layouts_production_capacity[size], qty
                        )
                    else:
                        layouts_production_capacity[size] = qty

        # Check if there are layouts that produce each demanded size
        missing_sizes = []
        insufficient_sizes = []
        for size, demand_qty in piece["quantity"].items():
            if size not in layouts_production_capacity:
                missing_sizes.append(size)
            elif layouts_production_capacity[size] == 0:
                insufficient_sizes.append(size)

        if missing_sizes:
            raise ValueError(
                f"No layouts can produce the following sizes: {missing_sizes}"
            )

        if insufficient_sizes:
            raise ValueError(
                f"Layouts have zero capacity for the following sizes: {insufficient_sizes}"
            )

        logging.info("Production capacity of layouts by size:")
        for size in sorted(layouts_production_capacity.keys()):
            logging.info(
                f"  {size}: maximum of {layouts_production_capacity[size]} pieces per layer"
            )

    def _validate_layout_sizes(self, layout):
        """Validate if the sizes in the layout are compatible with the extracted sizes.

        This method checks each piece in the layout to ensure that all sizes specified
        in the layout are valid and exist in the predefined set of sizes. It raises a
        ValueError if any invalid sizes are found.

        Args:
            layout (dict): The layout data containing:
                - pieces (list): List of pieces in the layout, each with a size grade.

        Raises:
            ValueError: If any sizes in the layout are invalid.
        """
        for piece in layout["pieces"]:
            layout_sizes = set(piece["size_grade"].keys())
            if not layout_sizes.issubset(set(self.sizes)):
                invalid_sizes = layout_sizes - set(self.sizes)
                raise ValueError(
                    f"Invalid sizes in layout {layout['id']}: {invalid_sizes}"
                )

    def _validate_solution(self, solution, demand):
        """Validate the optimization result.

        This method checks the validity of the optimization solution, ensuring that it
        is not None and that all required metrics are present. It raises a ValueError
        if any validation checks fail.

        Args:
            solution (dict): The optimization result containing metrics and production details.
            demand (dict): The demand data for validation purposes.

        Raises:
            ValueError: If the solution is invalid or if any required metrics are missing.
        """
        if not solution:
            raise ValueError("Invalid solution (None)")

        # Check if all required metrics exist
        required_metrics = [
            "fabric_meters",
            "fabric_cost",
            "cutting_cost",
            "layout_cost",
            "layer_cost",
            "waste_cost",
            "total_cost",
        ]

        for metric in required_metrics:
            if metric not in solution["metrics"]:
                raise ValueError(f"Missing metric: {metric}")

    def export_results(self, results):
        """Export results to an Excel file.

        This method creates an Excel workbook and populates it with the optimization
        results, including order details, layout information, and various metrics.

        Args:
            results (dict): A dictionary containing the results of the optimization,
                            indexed by order ID.

        Returns:
            None
        """
        wb = Workbook()
        ws = wb.active
        ws.title = "Results"

        # Adjusted headers for clarity
        headers = [
            "Order",
            "Pattern",
            "Layout",
            "Layers",
            "Total Length (m)",
            "Utilization (%)",
            "S",
            "M",
            "L",
            "XL",
            "Total Cost (R$)",
            "Total Fabric (m)",
            "Waste (m²)",
            "Waste (m)",
            "Waste (%)",
        ]
        ws.append(headers)

        # Data
        for order_id, result in results.items():
            if result:
                for layout in result["layouts_used"]:
                    row = [
                        order_id,
                        result["pattern"],
                        layout["layout_id"],
                        layout["num_layers"],
                        layout["length_meters"],
                        f"{layout['utilization']*100:.2f}",
                        layout["production_per_size"].get("P", 0),
                        layout["production_per_size"].get("M", 0),
                        layout["production_per_size"].get("G", 0),
                        layout["production_per_size"].get("GG", 0),
                        f"R$ {result['metrics']['total_cost']:.2f}",
                        f"{result['metrics']['fabric_meters']:.3f}",
                        f"{result['metrics']['fabric_waste_area']:.3f}",
                        f"{result['metrics']['fabric_waste_meters']:.3f}",
                        f"{result['metrics']['total_waste_percentage']:.2f}%",
                    ]
                    ws.append(row)

        wb.save("results.xlsx")

    @staticmethod
    def remove_accents(text):
        """Remove accents and special characters from a string.

        This method normalizes the input text to remove diacritics (accents) and
        special characters, returning a clean string. If an error occurs during
        processing, the original text is returned.

        Args:
            text (str): The input string from which accents and special characters
                        will be removed.

        Returns:
            str: The cleaned string without accents or special characters.
        """
        try:
            text = str(text)  # Ensure the input is a string
            normalized = unicodedata.normalize("NFKD", text)  # Normalize the string
            return "".join(
                [c for c in normalized if not unicodedata.combining(c)]
            )  # Remove accents
        except Exception:
            return text  # Return the original text if an error occurs

    def export_results_json(self, results):
        """Export results to a JSON file.

        This method processes the optimization results and exports them to a JSON file,
        including metrics, demand, production, and layout information with explicit units.

        Args:
            results (dict): A dictionary containing the results of the optimization,
                            indexed by order ID.

        Returns:
            None
        """
        try:
            output = []
            for order_id, result in results.items():
                if result:

                    def convert_variable(value):
                        """Convert solution variable to integer if it has a solution value."""
                        if hasattr(value, "solution_value"):
                            return int(value.solution_value())
                        return int(value)

                    # Metrics with explicit units
                    metrics = {
                        "fabric_meters": {
                            "value": round(
                                float(result["metrics"]["fabric_meters"]), 3
                            ),
                            "unit": "m",
                        },
                        "fabric_cost": {
                            "value": round(float(result["metrics"]["fabric_cost"]), 2),
                            "unit": "BRL",
                        },
                        "cutting_cost": {
                            "value": round(float(result["metrics"]["cutting_cost"]), 2),
                            "unit": "BRL",
                        },
                        "layout_cost": {
                            "value": round(float(result["metrics"]["layout_cost"]), 2),
                            "unit": "BRL",
                        },
                        "layer_cost": {
                            "value": round(float(result["metrics"]["layer_cost"]), 2),
                            "unit": "BRL",
                        },
                        "waste_cost": {
                            "value": round(float(result["metrics"]["waste_cost"]), 2),
                            "unit": "BRL",
                        },
                        "total_cost": {
                            "value": round(
                                sum(
                                    [
                                        float(result["metrics"]["fabric_cost"]),
                                        float(result["metrics"]["cutting_cost"]),
                                        float(result["metrics"]["layout_cost"]),
                                        float(result["metrics"]["layer_cost"]),
                                        float(result["metrics"]["waste_cost"]),
                                    ]
                                ),
                                2,
                            ),
                            "unit": "BRL",
                        },
                        "fabric_waste_area": {
                            "value": round(
                                float(result["metrics"]["fabric_waste_area"]), 4
                            ),
                            "unit": "m²",
                        },
                        "fabric_waste_meters": {
                            "value": round(
                                float(result["metrics"]["fabric_waste_meters"]), 4
                            ),
                            "unit": "m",
                        },
                        "total_waste_percentage": {
                            "value": round(
                                float(result["metrics"]["total_waste_percentage"]), 2
                            ),
                            "unit": "%",
                        },
                    }

                    output_item = {
                        "order_id": order_id,
                        "pattern": result["pattern"],
                        "metrics": metrics,
                        "demand": {
                            size: convert_variable(qty)
                            for size, qty in result.get("demand", {}).items()
                        },
                        "production": {
                            size: convert_variable(qty)
                            for size, qty in result.get("production", {}).items()
                        },
                        "overproduction": {
                            size: convert_variable(qty)
                            for size, qty in result.get("overproduction", {}).items()
                        },
                        "layouts_used": [],
                    }

                    # Process layouts with explicit units
                    for layout in result.get("layouts_used", []):
                        layout_info = {
                            "layout_id": layout["layout_id"],
                            "num_layers": {
                                "value": convert_variable(layout["num_layers"]),
                                "unit": "layers",
                            },
                            "length_meters": {
                                "value": round(float(layout["length_meters"]), 3),
                                "unit": "m",
                            },
                            "utilization": {
                                "value": round(float(layout["utilization"]), 4),
                                "unit": "%",
                            },
                            "waste_area": {
                                "value": round(float(layout["waste_area"]), 4),
                                "unit": "m²",
                            },
                            "production_per_size": {
                                size: {"value": convert_variable(qty), "unit": "pieces"}
                                for size, qty in layout["production_per_size"].items()
                            },
                            "costs": {
                                "fabric_cost": {
                                    "value": round(
                                        float(layout["costs"]["fabric_cost"]), 2
                                    ),
                                    "unit": "BRL",
                                },
                                "cutting_cost": {
                                    "value": round(
                                        float(layout["costs"]["cutting_cost"]), 2
                                    ),
                                    "unit": "BRL",
                                },
                                "layout_cost": {
                                    "value": round(
                                        float(layout["costs"]["layout_cost"]), 2
                                    ),
                                    "unit": "BRL",
                                },
                            },
                        }
                        output_item["layouts_used"].append(layout_info)

                    output.append(output_item)

            # Write the result to a JSON file
            with open("results.json", "w", encoding="utf-8") as f:
                json.dump(output, f, indent=4, ensure_ascii=False)

        except Exception as e:
            logging.error(f"Error exporting results to JSON: {str(e)}")
            logging.error(traceback.format_exc())
            raise


def main():
    """Main function of the optimizer.

    This function loads input data, processes each piece as a separate order,
    optimizes the production layout, and exports the results to both Excel and JSON files.

    Returns:
        None
    """
    try:
        # Load input data
        with open("input_data.json", "r", encoding="utf-8") as f:
            input_data = json.load(f)

        logging.info("Starting production optimization")
        optimizer = LayoutOptimizer(input_data)

        # Process each piece as a separate order
        results = {}
        for i, piece in enumerate(optimizer.pieces, 1):
            order_id = f"order_{i}"
            logging.info(f"Processing {piece['pattern']} (Order {order_id})")

            # Create order structure
            order = {
                "id": order_id,
                "pieces": [piece],
                "fabric_width": optimizer.fabrics[piece["fabrics"][0]]["fabric_width"],
                "max_layers": optimizer.fabrics[piece["fabrics"][0]]["max_layers"],
                "max_length": optimizer.config["max_total_length"],
            }

            # Optimize order
            result = optimizer.optimize_order(order)
            results[order_id] = result

            # Log results
            if result:
                logging.info(f"Solution found for {piece['pattern']}")
                logging.info("Production by size:")
                for size in optimizer.sizes:
                    prod = result["production"][size]
                    over = result["overproduction"][size]
                    logging.info(f"  {size}: {prod} (Excess: {over})")

                metrics = result["metrics"]
                logging.info(f"Total cost: R$ {metrics['total_cost']:.2f}")
                logging.info(f"Fabric meters: {metrics['fabric_meters']:.2f} m")
                logging.info(f"Waste: {metrics['fabric_waste_area']:.4f} m²")
            else:
                logging.warning(f"Could not find a solution for {piece['pattern']}")

        # Export results
        if any(results.values()):
            optimizer.export_results(results)
            optimizer.export_results_json(results)
            logging.info("Results exported to results.xlsx and results.json")
        else:
            logging.warning("No solutions found")

    except Exception as e:
        logging.error(f"Execution error: {str(e)}")
        raise


if __name__ == "__main__":
    main()
