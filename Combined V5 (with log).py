import pandas as pd
import tkinter as tk
from tkinter import simpledialog, filedialog
from ib_insync import IB
from ibapi.client import EClient
from ibapi.wrapper import EWrapper
from ibapi.common import *
from ibapi.contract import Contract
import logging
import time

# Set up logging to capture any issues and save to file
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', handlers=[
    logging.FileHandler("ib_script.log", mode='w', encoding='utf-8'),
    logging.StreamHandler()
])

log_messages = []

def log_and_store(message, level=logging.INFO):
    # Log to the logger and also store in log_messages
    logging.log(level, message)
    log_messages.append(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {message}")
    # Flush the logger to ensure all logs are written immediately
    logging.getLogger().handlers[0].flush()

# Updated Script 1: Account Summary with Net Liquidation and Total Cash using IBInsync
def fetch_ib_account_data():
    ib = IB()
    ib.connect('127.0.0.1', 7496, clientId=1)
    
    account_values = ib.accountValues()
    ib.disconnect()
    
    data = []
    
    for value in account_values:
        account = value.account
        tag = value.tag
        value_amount = value.value
        currency = value.currency
        
        if tag == "NetLiquidation":
            data.append([account, "Net Liquidation Value", float(value_amount), currency])
        elif tag == "TotalCashBalance" and account != "All":
            data.append([account, "Total Cash Balance", float(value_amount), currency])
    
    return pd.DataFrame(data, columns=["Account", "Tag", "Total Value", "Base Currency"])

# Script 2: Live Positions
class IBApp(EClient, EWrapper):
    def __init__(self):
        EClient.__init__(self, self)
        self.positions = []

    def position(self, account, contract, position, avgCost):
        contract_details = {
            "Symbol": contract.symbol,
            "SecType": contract.secType,
            "Exchange": contract.exchange,
            "Currency": contract.currency,
            "Strike": contract.strike if hasattr(contract, "strike") else "N/A",
            "Expiry": contract.lastTradeDateOrContractMonth if hasattr(contract, "lastTradeDateOrContractMonth") else "N/A",
            "Right": contract.right if hasattr(contract, "right") else "N/A",
            "Mult": contract.multiplier if hasattr(contract, "multiplier") else "N/A",
        }
        
        log_and_store(f"Position received: {contract_details}", logging.DEBUG)
        
        self.positions.append({
            "Account": account,
            "Symbol": contract.symbol,
            "Position": position,
            "Avg Cost": avgCost,
            **contract_details
        })

    def positionEnd(self):
        log_and_store("Position End reached.")
        
        if not self.positions:
            log_and_store("No positions to save.", logging.WARNING)
            return
        
        df_positions = pd.DataFrame(self.positions)
        log_and_store(f"Positions DataFrame created: {df_positions}", logging.DEBUG)
        
        root = tk.Tk()
        root.withdraw()
        
        file_name = simpledialog.askstring("Input", "Enter file name (without extension):")
        if not file_name:
            log_and_store("Save canceled by user.")
            return
        
        file_path = filedialog.asksaveasfilename(
            initialfile=f"{file_name}.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Save Excel File"
        )
        
        if not file_path:
            log_and_store("Save canceled by user.")
            return
        
        try:
            df_summary = fetch_ib_account_data()
            
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                if not df_summary.empty:
                    df_summary.to_excel(writer, sheet_name="Account Summary", index=False)
                if not df_positions.empty:
                    df_positions.to_excel(writer, sheet_name="Positions", index=False)
                
            log_and_store(f"Data saved to {file_path}")
        
        except Exception as e:
            log_and_store(f"Error saving file: {e}", logging.ERROR)
        
        finally:
            root.destroy()
        
        # Save the full log
        log_file_path = filedialog.asksaveasfilename(
            initialfile=f"{file_name}_log.txt",
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt")],
            title="Save Log File"
        )
        
        if log_file_path:
            try:
                with open(log_file_path, "w", encoding="utf-8") as log_file:
                    # Write the full log
                    log_file.write("\n".join(log_messages))
                log_and_store(f"Log saved to {log_file_path}")
            except Exception as e:
                log_and_store(f"Error saving log: {e}", logging.ERROR)
        
        log_and_store("Waiting for 1 second before disconnecting...")
        time.sleep(1)
        
        log_and_store("Disconnecting from IB...")
        self.disconnect()
        print("Disconnected and script finished.")

    def managedAccounts(self, accountsList):
        log_and_store(f"Managed Accounts: {accountsList}")

    def run(self):
        log_and_store("Connecting to IB API.")
        self.connect("127.0.0.1", 7496, 0)
        self.reqManagedAccts()
        self.reqPositions()
        log_and_store("Starting event loop.")
        super().run()

if __name__ == "__main__":
    try:
        app = IBApp()
        app.run()
    except Exception as e:
        log_and_store(f"Error in script execution: {e}", logging.ERROR)
        print(f"Error: {e}")
    
    input("Press Enter to exit...")
