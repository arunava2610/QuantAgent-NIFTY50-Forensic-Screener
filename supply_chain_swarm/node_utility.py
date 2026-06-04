import pandas as pd
import requests
import sys
import os

# Import config from parent directory
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import logger, ERP_DRIVE_URL

def run_node_allocation(missing_item: str):
    """Downloads ERP data and scans for replacement inventory."""
    temp_file = "temp_erp_matrix.xlsx"

    try:
        # 1. Securely fetch the live database from Drive
        headers = {'User-Agent': 'Mozilla/5.0'}
        res = requests.get(ERP_DRIVE_URL, headers=headers)
        if res.status_code == 200:
            with open(temp_file, "wb") as f:
                f.write(res.content)
        else:
            raise Exception(f"Drive connection failed with status: {res.status_code}")

        # 2. Analyze the Inventory Nodes
        df = pd.read_excel(temp_file, sheet_name="Inventory_Nodes") # type: ignore

        # 3. Search for the item (Using fuzzy matching in case AI shortened the name)
        # Looking for matching name OR SKU where quantity is greater than zero
        match = df[
            (df['Item_Name'].str.contains(missing_item, case=False, na=False) | 
             df['SKU'].str.contains(missing_item, case=False, na=False)) & 
            (df['Quantity_Available'] > 0)
        ]

        if match.empty:
            return None, 0

        # 4. Find the best node (Sorting by highest available stock to be safe)
        best_node = match.sort_values(by='Quantity_Available', ascending=False).iloc[0]
        node_name = best_node['Location']
        available_qty = int(best_node['Quantity_Available'])

        # Reserve a standard operational chunk (e.g., 500 units, or whatever is left)
        reserve_amount = min(500, available_qty)

        return node_name, reserve_amount

    except Exception as e:
        logger.error(f"Node Allocation Error: {str(e)}")
        raise e