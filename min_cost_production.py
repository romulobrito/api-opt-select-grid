import logging
from ortools.linear_solver import pywraplp
import random
import datetime
import plotly.figure_factory as ff
import numpy as np
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import plotly.io as pio
import os

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("debug.log"), logging.StreamHandler()],
)
SIZES = ["P", "M", "G", "GG"] 
class Resource:
    def __init__(self, id, efficiency):
        self.id = id
        self.efficiency = efficiency

class Shift:
    def __init__(self, start, end, efficiency=1.0):
        self.start = start
        self.end = end
        self.efficiency = efficiency

def save_html_figure(fig, file_path="gantt_chart.html"):
    try:
        fig.write_html(file_path)
        logging.info(f"HTML figure saved to {file_path}")
    except Exception as e:
        logging.error(f"Error saving HTML figure: {e}")

def export_to_excel(
    schedule,
    ordered_orders,
    results,
    orders,
    sizes,
    layouts,
    fig,
    filename=r"G:\My Drive\senai_sc\production_results.xlsx",
):
    wb = Workbook()
    ws = wb.active
    ws.title = "Production Schedule"

    # Header styles
    header_font = Font(bold=True)
    header_fill = PatternFill(
        start_color="DDDDDD", end_color="DDDDDD", fill_type="solid"
    )

    # Schedule
    ws.cell(row=1, column=1, value="Production Schedule").font = Font(
        bold=True, size=14
    )
    headers = ["Order", "Operation", "Start", "End", "Resource"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill

    row = 4
    for p in ordered_orders:
        ws.cell(row=row, column=1, value=f"Order {p}")
        ws.cell(row=row, column=2, value="Spreading")
        ws.cell(row=row, column=3, value=schedule[p]["spreading_start"])
        ws.cell(row=row, column=4, value=schedule[p]["spreading_end"])
        ws.cell(
            row=row, column=5, value=f"Spreader {schedule[p]['spreader']}"
        )
        row += 1

        ws.cell(row=row, column=1, value=f"Order {p}")
        ws.cell(row=row, column=2, value="Cutting")
        ws.cell(row=row, column=3, value=schedule[p]["cutting_start"])
        ws.cell(row=row, column=4, value=schedule[p]["cutting_end"])
        ws.cell(
            row=row,
            column=5,
            value=f"Cutting Machine {schedule[p]['cutting_machine']}",
        )
        row += 1

    # Order details
    row += 2
    ws.cell(row=row, column=1, value="Order Details").font = Font(
        bold=True, size=14
    )
    row += 1
    headers = [
        "Order",
        "Deadline",
        "Total Cost",
        "Setup Cost",
        "Fabric Meters",
        "Cut Perimeter",
        "Waste",
        "Actual Spread Length",
        "Maximum Spread Length",
        "Relaxation",
    ]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
    row += 1

    for p in ordered_orders:
        result = results[p]
        ws.cell(row=row, column=1, value=f"Order {p}")
        ws.cell(row=row, column=2, value=orders[p]["deadline"])
        ws.cell(row=row, column=3, value=result["total_cost"])
        ws.cell(row=row, column=4, value=result["setup_cost"])
        ws.cell(row=row, column=5, value=result["fabric_meters"])
        ws.cell(row=row, column=6, value=result["cut_perimeter"])
        ws.cell(row=row, column=7, value=result["waste"])
        ws.cell(row=row, column=8, value=result["actual_spread_length"])
        ws.cell(row=row, column=9, value=result["max_spread_length"])
        ws.cell(row=row, column=10, value=str(result["relaxation"]))
        row += 1

    # Production details
    row += 2
    ws.cell(row=row, column=1, value="Production Details").font = Font(
        bold=True, size=14
    )
    row += 1

    headers = ["Order"] + sizes
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
    row += 1

    for p in ordered_orders:
        ws.cell(row=row, column=1, value=f"Order {p}")
        for col, size in enumerate(sizes, start=2):
            ws.cell(row=row, column=col, value=results[p]["production"][size])
        row += 1

    # Adjust column widths
    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 15

    try:
        wb.save(filename)
        logging.info(f"Excel file saved to {filename}")
    except Exception as e:
        logging.error(f"Error saving Excel file: {e}")

def export_orders_demand_excel(orders, filename="orders_demand.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Orders Demand"

    # Headers
    headers = ["Order", "Deadline"] + list(next(iter(orders.values()))["demands"].keys())
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=header)

    # Data
    for row, (order_id, order) in enumerate(orders.items(), start=2):
        ws.cell(row=row, column=1, value=order_id)
        ws.cell(row=row, column=2, value=order["deadline"].strftime("%Y-%m-%d"))
        for col, qty in enumerate(order["demands"].values(), start=3):
            ws.cell(row=row, column=col, value=qty)

    try:
        wb.save(filename)
        logging.info(f"Orders demand exported to {filename}")
    except Exception as e:
        logging.error(f"Error exporting orders demand: {e}")

def create_gantt_chart(schedule, ordered_orders, priority_criterion):
    tasks = []
    for order in ordered_orders:
        # Spreading task
        tasks.append({
            "Task": f"Order {order}",
            "Start": schedule[order]["spreading_start"],
            "Finish": schedule[order]["spreading_end"],
            "Resource": f"Spreader {schedule[order]['spreader']}"
        })
        # Cutting task
        tasks.append({
            "Task": f"Order {order}",
            "Start": schedule[order]["cutting_start"],
            "Finish": schedule[order]["cutting_end"],
            "Resource": f"Cutting Machine {schedule[order]['cutting_machine']}"
        })

    colors = {
        "Spreader 1": "rgb(46, 137, 205)",
        "Spreader 2": "rgb(114, 44, 121)",
        "Spreader 3": "rgb(198, 47, 105)",
        "Cutting Machine 1": "rgb(58, 149, 136)",
        "Cutting Machine 2": "rgb(107, 127, 135)"
    }

    fig = ff.create_gantt(
        tasks,
        colors=colors,
        index_col="Resource",
        title=f"Production Schedule (Priority: {priority_criterion})",
        show_colorbar=True,
        group_tasks=True,
        showgrid_x=True,
        showgrid_y=True
    )

    return fig

def calculate_production_time(result, layouts, resources):
    return result["spreading_time"] + result["cutting_time"]

def generate_schedule(ordered_orders, results, layouts, start_date, resources, shifts):
    schedule = {}
    current_time = {
        "spreaders": {r.id: start_date for r in resources["spreaders"]},
        "cutting_machines": {r.id: start_date for r in resources["cutting_machines"]}
    }

    for order in ordered_orders:
        result = results[order]
        
        # Find earliest available spreader
        spreader_id = min(
            current_time["spreaders"],
            key=lambda x: current_time["spreaders"][x]
        )
        
        spreading_start = current_time["spreaders"][spreader_id]
        spreading_duration = datetime.timedelta(hours=result["spreading_time"])
        spreading_end = spreading_start + spreading_duration
        
        # Update spreader availability
        current_time["spreaders"][spreader_id] = spreading_end
        
        # Find earliest available cutting machine
        cutting_machine_id = min(
            current_time["cutting_machines"],
            key=lambda x: current_time["cutting_machines"][x]
        )
        
        cutting_start = max(spreading_end, current_time["cutting_machines"][cutting_machine_id])
        cutting_duration = datetime.timedelta(hours=result["cutting_time"])
        cutting_end = cutting_start + cutting_duration
        
        # Update cutting machine availability
        current_time["cutting_machines"][cutting_machine_id] = cutting_end
        
        schedule[order] = {
            "spreading_start": spreading_start,
            "spreading_end": spreading_end,
            "spreader": spreader_id,
            "cutting_start": cutting_start,
            "cutting_end": cutting_end,
            "cutting_machine": cutting_machine_id
        }
    
    return schedule



def optimize_order(
    order,
    layouts,
    width_tolerance,
    overproduction_percentage,
    max_layers_per_layout,
    production_hours,
    spread_table_length,
    overproduction_penalty,
    relaxation=False
):
    """
    Optimizes production for a single order using linear programming.
    
    Args:
        order: Dictionary containing order details (demands, deadline)
        layouts: Dictionary of available layouts with their specifications
        width_tolerance: Maximum allowed width tolerance
        overproduction_percentage: Maximum allowed overproduction percentage
        max_layers_per_layout: Maximum number of layers per layout
        production_hours: Available production hours
        spread_table_length: Length of spreading table
        overproduction_penalty: Penalty cost for overproduction
        relaxation: Whether to use relaxation in the solver
    
    Returns:
        Dictionary containing optimization results
    """
    solver = pywraplp.Solver.CreateSolver('SCIP')
    if not solver:
        raise ValueError('Could not create solver')

    # Create variables for each layout
    layout_vars = {}
    for layout_id, layout in layouts.items():
        layout_vars[layout_id] = solver.IntVar(0, max_layers_per_layout, f'layout_{layout_id}')

    # Create variables for production quantities
    production_vars = {}
    for size in order['demands'].keys():
        production_vars[size] = solver.NumVar(0, solver.infinity(), f'production_{size}')

    # Create variables for overproduction
    overproduction_vars = {}
    for size in order['demands'].keys():
        overproduction_vars[size] = solver.NumVar(0, solver.infinity(), f'over_{size}')

    # Demand constraints
    for size, demand in order['demands'].items():
        solver.Add(
            production_vars[size] - overproduction_vars[size] == demand,
            f'demand_{size}'
        )

    # Production constraints
    for size in order['demands'].keys():
        size_production = solver.Sum(
            layout_vars[layout_id] * layout['quantities'][size]
            for layout_id, layout in layouts.items()
        )
        solver.Add(production_vars[size] == size_production, f'production_{size}')

    # Overproduction constraints
    for size, demand in order['demands'].items():
        solver.Add(
            overproduction_vars[size] <= demand * overproduction_percentage,
            f'max_over_{size}'
        )

    # Calculate total cost
    total_cost = solver.Sum([
        layout_vars[layout_id] * (
            layout['setup_cost'] +
            layout['fabric_cost'] * layout['fabric_length'] +
            layout['cutting_cost'] * layout['total_perimeter']
        )
        for layout_id, layout in layouts.items()
    ])

    # Add overproduction penalty to total cost
    overproduction_cost = solver.Sum([
        overproduction_vars[size] * overproduction_penalty
        for size in order['demands'].keys()
    ])

    total_cost += overproduction_cost

    # Set objective
    solver.Minimize(total_cost)

    # Solve
    status = solver.Solve()

    if status == pywraplp.Solver.OPTIMAL:
        result = {
            'status': 'optimal',
            'total_cost': solver.Objective().Value(),
            'production': {
                size: production_vars[size].solution_value()
                for size in order['demands'].keys()
            },
            'layers': {
                layout_id: layout_vars[layout_id].solution_value()
                for layout_id in layouts.keys()
            },
            'overproduction': {
                size: overproduction_vars[size].solution_value()
                for size in order['demands'].keys()
            },
            'metrics': calculate_metrics(
                layouts,
                {layout_id: layout_vars[layout_id].solution_value() for layout_id in layouts.keys()}
            )
        }
        return result
    else:
        logging.warning(f"No optimal solution found. Status: {status}")
        return None

def calculate_metrics(layouts, used_layers):
    """
    Calculates production metrics based on the optimization results.
    
    Args:
        layouts: Dictionary of layouts with their specifications
        used_layers: Dictionary of number of layers used for each layout
    
    Returns:
        Dictionary containing calculated metrics
    """
    fabric_meters = sum(
        layout['fabric_length'] * layers
        for layout_id, (layout, layers) in zip(layouts.keys(), zip(layouts.values(), used_layers.values()))
    )
    
    cut_perimeter = sum(
        layout['total_perimeter'] * layers
        for layout_id, (layout, layers) in zip(layouts.keys(), zip(layouts.values(), used_layers.values()))
    )
    
    setup_cost = sum(
        layout['setup_cost'] * (layers > 0)
        for layout_id, (layout, layers) in zip(layouts.keys(), zip(layouts.values(), used_layers.values()))
    )
    
    total_cost = setup_cost + sum(
        (layout['fabric_cost'] * layout['fabric_length'] +
         layout['cutting_cost'] * layout['total_perimeter']) * layers
        for layout_id, (layout, layers) in zip(layouts.keys(), zip(layouts.values(), used_layers.values()))
    )
    
    return {
        'fabric_meters': fabric_meters,
        'cut_perimeter': cut_perimeter,
        'setup_cost': setup_cost,
        'total_cost': total_cost
    }


def get_sizes_from_orders(orders):
    """Extract unique sizes from orders"""
    sizes = set()
    for order in orders.values():
        sizes.update(order['demands'].keys())
    return sorted(list(sizes))

def get_sizes_from_layouts(layouts):
    """Extract unique sizes from layouts"""
    sizes = set()
    for layout in layouts.values():
        sizes.update(layout['quantities'].keys())
    return sorted(list(sizes))

def main(
    priority_criterion="deadline",
    start_date=None,
    num_days=5,
    resources=None,
    width_tolerance=0.05,
    overproduction_percentage=0.05,
    max_layers_per_layout=30,
    layouts=None,
    orders=None,
    shifts=None,
    production_hours=24,
    spread_table_length=10.0,
    overproduction_penalty=10000,
    relaxation=False,
):
    if start_date is None:
        start_date = datetime.datetime.now()

    sizes_from_orders = get_sizes_from_orders(orders) if orders else SIZES
    sizes_from_layouts = get_sizes_from_layouts(layouts) if layouts else SIZES

    sizes = sorted(list(set(sizes_from_orders) & set(sizes_from_layouts)))
    
    if not sizes:
        raise ValueError("No matching sizes found between orders and layouts")

    # Convert resources dict to objects
    resources_obj = {
        "spreaders": [Resource(r["id"], r["efficiency"]) for r in resources["spreaders"]],
        "cutting_machines": [Resource(r["id"], r["efficiency"]) for r in resources["cutting_machines"]]
    }

    # Initialize results dictionary
    results = {}
    
    # Process each order
    for order_id, order in orders.items():
        result = optimize_order(
            order,
            layouts,
            width_tolerance,
            overproduction_percentage,
            max_layers_per_layout,
            production_hours,
            spread_table_length,
            overproduction_penalty,
            relaxation
        )
        results[order_id] = result

    def calculate_priority(order, result, criterion):
        if criterion == "deadline":
            return (order["deadline"] - start_date).days
        elif criterion == "total_cost":
            return result["total_cost"]
        elif criterion == "production_time":
            return calculate_production_time(result, layouts, resources_obj)
        else:
            raise ValueError(f"Invalid priority criterion: {criterion}")

    priorities = {
        p: calculate_priority(orders[p], result, priority_criterion)
        for p, result in results.items()
    }

    ordered_orders = sorted(priorities, key=priorities.get)

    schedule = generate_schedule(
        ordered_orders, results, layouts, start_date, resources_obj, shifts
    )

    # Display results and chart
    print(f"\nOrder prioritization ({priority_criterion}):")
    for i, p in enumerate(ordered_orders, 1):
        print(f"{i}. Order {p}: {priorities[p]:.2f}")

    print("\nProduction schedule:")
    for p in ordered_orders:
        print(f"Order {p}:")
        print(
            f"  Spreading: Start {schedule[p]['spreading_start'].strftime('%d/%m/%Y %H:%M')} - "
            f"End {schedule[p]['spreading_end'].strftime('%d/%m/%Y %H:%M')} "
            f"(Spreader {schedule[p]['spreader']})"
        )
        print(
            f"  Cutting: Start {schedule[p]['cutting_start'].strftime('%d/%m/%Y %H:%M')} - "
            f"End {schedule[p]['cutting_end'].strftime('%d/%m/%Y %H:%M')} "
            f"(Cutting Machine {schedule[p]['cutting_machine']})"
        )

    print("\nOrder details:")
    for p in ordered_orders:
        result = results[p]
        print(f"\nOrder {p}:")
        print(f"  Deadline: {orders[p]['deadline'].strftime('%d/%m/%Y')}")
        print(f"  Total cost: $ {result['total_cost']:.2f}")
        print(f"  Fabric meters: {result['fabric_meters']:.2f} m")
        print(f"  Cut perimeter: {result['cut_perimeter']:.2f} m")
        print(f"  Fabric waste: {result['waste']:.2f} m²")
        print(f"  Actual spreading time: {result['spreading_time']:.2f} h")
        print(f"  Actual cutting time: {result['cutting_time']:.2f} h")
        print(f"  Total actual time: {result['total_time']:.2f} h")
        print("  Production:")
        for size in sizes:
            print(
                f"    {size}: {int(result['production'][size])} pieces (Demand: {orders[p]['demands'][size]})"
            )
        print("  Layers used:")
        for layout, layers in result["layers"].items():
            if layers > 0:
                print(f"    {layout}: {int(layers)} layers")

        print(f"  Setup cost: $ {result['setup_cost']:.2f}")

    fig = create_gantt_chart(schedule, ordered_orders, priority_criterion)
    pio.show(fig)

    export_to_excel(
        schedule, ordered_orders, results, orders, sizes, layouts, fig
    )
    export_orders_demand_excel(
        orders, f"orders_demand-{priority_criterion}.xlsx"
    )
    save_html_figure(fig, f"gantt_chart-{priority_criterion}.html")

    # Prepare detailed results
    detailed_results = {}
    for p in ordered_orders:
        result = results[p]
        detailed_results[p] = {
            "deadline": orders[p]["deadline"].isoformat(),
            "total_cost": float(result["total_cost"]),
            "fabric_meters": float(result["fabric_meters"]),
            "cut_perimeter": float(result["cut_perimeter"]),
            "waste": float(result["waste"]),
            "spreading_time": float(result["spreading_time"]),
            "cutting_time": float(result["cutting_time"]),
            "total_time": float(result["total_time"]),
            "production": {size: int(result["production"][size]) for size in sizes},
            "demands": orders[p]["demands"],
            "layers": {layout: int(layers) for layout, layers in result["layers"].items() if layers > 0},
            "setup_cost": float(result["setup_cost"])
        }

    # Return results
    return {
        "ordered_orders": ordered_orders,
        "schedule": {
            str(k): {
                "spreading_start": v["spreading_start"].isoformat(),
                "spreading_end": v["spreading_end"].isoformat(),
                "spreader": v["spreader"],
                "cutting_start": v["cutting_start"].isoformat(),
                "cutting_end": v["cutting_end"].isoformat(),
                "cutting_machine": v["cutting_machine"]
            }
            for k, v in schedule.items()
        },
        "results": detailed_results,
        "global_metrics": {
            "total_cost": sum(r["total_cost"] for r in detailed_results.values()),
            "total_time": sum(r["total_time"] for r in detailed_results.values()),
            "total_waste": sum(r["waste"] for r in detailed_results.values())
        }
    }

if __name__ == "__main__":
    main()