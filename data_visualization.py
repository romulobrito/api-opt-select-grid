import requests
import json
from datetime import datetime


def format_date(date_str):
    """Format a date string to a specific format.

    Args:
        date_str (str): The date string in ISO format.

    Returns:
        str: The formatted date string in "dd/mm/yyyy HH:MM" format.
    """
    return datetime.fromisoformat(date_str).strftime("%d/%m/%Y %H:%M")


def format_currency(value):
    """Format a numeric value as currency.

    Args:
        value (float): The numeric value to format.

    Returns:
        str: The formatted currency string.
    """
    return f"R$ {value:,.2f}"


def get_token():
    """Function to obtain authentication token.

    Returns:
        str: The access token if successful, None otherwise.
    """
    login_url = "http://localhost:8000/token"

    data = {"username": "admin", "password": "admin123", "grant_type": "password"}

    try:
        response = requests.post(
            login_url,
            data=data,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
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
    """Main function to view optimization results.

    This function retrieves the optimization results from the API and displays
    them in a formatted manner.

    Returns:
        None
    """
    print("Starting visualization script...")
    token = get_token()
    if not token:
        print("Could not obtain authentication token")
        return

    url = "http://localhost:8000/optimize"

    try:
        with open("input_data.json", "r", encoding="utf-8") as f:
            data = json.load(f)

        # Basic data validation
        required_keys = ["general_configuration", "layouts", "fabrics", "pieces"]
        for key in required_keys:
            if key not in data:
                raise ValueError(f"Missing required key: {key}")

        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

        print(f"Sending request to {url}...")
        response = requests.post(url, json=data, headers=headers)
        print(f"Response status: {response.status_code}")

        if response.status_code == 422:
            error_detail = response.json().get("detail", "Unknown validation error")
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
                    print(f"\nItem: {result.get('pattern', 'Unknown')}")

                    print("\nDemand vs Production:")
                    for size in result["production"].keys():
                        demand = result["demand"].get(size, 0)
                        prod = result["production"].get(size, 0)
                        over = result["overproduction"].get(size, 0)
                        print(
                            f"  {size}: Demand={demand}, Production={prod} (Excess: {over})"
                        )

                    metrics = result["metrics"]
                    print("\nWaste Metrics:")
                    print(f"  Waste Area: {metrics['fabric_waste_area']:.4f} m²")
                    print(f"  Waste Length: {metrics['fabric_waste_meters']:.4f} m")
                    print(
                        f"  Waste Percentage: {metrics['total_waste_percentage']:.1f}%"
                    )

                    print("\nCosts (KPIs):")
                    try:
                        print(
                            f"  Fabric Cost: {format_currency(metrics['fabric_cost'])}"
                        )
                        print(
                            f"  Cutting Cost: {format_currency(metrics['cutting_cost'])}"
                        )
                        print(
                            f"  Setup Cost: {format_currency(metrics['layout_setup_cost'])}"
                        )
                        print(
                            f"  Cost per Layer: {format_currency(metrics['layer_cost'])}"
                        )
                        print(f"  Waste Cost: {format_currency(metrics['waste_cost'])}")
                        print(f"  Total Cost: {format_currency(metrics['total_cost'])}")
                    except KeyError as e:
                        print(f"Error: Missing metric: {e}")

                    print("\nLayouts Used:")
                    for layout in result.get("layouts_used", []):
                        print(f"\n  Layout {layout['layout_id']}:")
                        print(f"    Layers: {layout['num_layers']}")
                        print(f"    Length: {layout['length_meters']:.3f} m")
                        print(f"    Utilization: {layout['utilization']*100:.1f}%")
                        print(f"    Waste: {layout['waste_area']:.4f} m²")
                else:
                    print(f"\nItem {order_id}: No viable solution")

    except requests.exceptions.ConnectionError:
        print("\nError: Could not connect to the API.")
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        import traceback

        print(traceback.format_exc())


if __name__ == "__main__":
    view_results()
