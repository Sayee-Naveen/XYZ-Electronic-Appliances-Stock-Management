import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
import os

def create_sample_excel(file_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Appliance_Stock"

    # Define headers
    headers = [
        "SKU", "Product_Name", "Category", "Brand", "Model_Number",
        "Stock_Quantity", "Min_Threshold", "Cost_Price", "Selling_Price",
        "Warehouse_Location", "Energy_Rating", "Warranty", "Supplier",
        "Status", "Last_Updated"
    ]

    # Sample realistic data for XYZ Electronic Appliances
    data = [
        ["XYZ-TV-001", "Sony Bravia 65-Inch 4K Ultra HD Smart LED TV", "Televisions", "Sony", "KD-65X74L", 18, 5, 58000, 74990, "Bay A-10", "4 Star", "2 Years", "Pacific Tech Supplies", "In Stock", "2026-11-20 10:15"],
        ["XYZ-TV-002", "Samsung 55-Inch Crystal 4K Vivid Pro Smart TV", "Televisions", "Samsung", "UA55DUE77AK", 4, 6, 32000, 43990, "Bay A-11", "3 Star", "1 Year", "Apex Electronics Distro", "Low Stock", "2026-11-20 11:30"],
        ["XYZ-TV-003", "LG OLED evo C3 48-Inch 4K Smart Cinema TV", "Televisions", "LG", "OLED48C3PSA", 0, 3, 72000, 89990, "Bay A-12", "5 Star", "3 Years", "Apex Electronics Distro", "Out of Stock", "2026-11-19 16:45"],
        ["XYZ-REF-101", "Samsung 253L 3-Star Inverter Double Door Refrigerator", "Refrigerators", "Samsung", "RT28C3733S8", 22, 8, 21000, 26990, "Bay B-01", "3 Star", "10 Yrs Compressor", "Metro Appliance Wholesale", "In Stock", "2026-11-20 09:20"],
        ["XYZ-REF-102", "Whirlpool 265L 2-Star Frost Free Double Door Refrigerator", "Refrigerators", "Whirlpool", "IF INV CNV 278", 5, 7, 22500, 28490, "Bay B-02", "2 Star", "10 Yrs Compressor", "Metro Appliance Wholesale", "Low Stock", "2026-11-20 14:10"],
        ["XYZ-REF-103", "LG 655L Frost-Free Smart Inverter Side-by-Side Refrigerator", "Refrigerators", "LG", "GL-B257HDSY", 8, 4, 65000, 79990, "Bay B-03", "4 Star", "10 Yrs Compressor", "Pacific Tech Supplies", "In Stock", "2026-11-18 12:00"],
        ["XYZ-REF-104", "Haier 190L 4-Star Single Door Direct Cool Refrigerator", "Refrigerators", "Haier", "HED-204DS-P", 14, 5, 11500, 14990, "Bay B-04", "4 Star", "5 Yrs Compressor", "Metro Appliance Wholesale", "In Stock", "2026-11-20 08:45"],
        ["XYZ-WM-201", "LG 8.0 Kg 5-Star AI Direct Drive Front Load Washing Machine", "Washing Machines", "LG", "FHP1208Z5W", 12, 5, 31000, 38990, "Bay C-01", "5 Star", "2 Years / 10 Yrs Motor", "Apex Electronics Distro", "In Stock", "2026-11-20 13:00"],
        ["XYZ-WM-202", "Bosch 7.0 Kg 5-Star Inverter Touch Front Load Washer", "Washing Machines", "Bosch", "WAJ2416SIN", 3, 5, 26000, 33490, "Bay C-02", "5 Star", "3 Years", "Pacific Tech Supplies", "Low Stock", "2026-11-19 18:20"],
        ["XYZ-WM-203", "Samsung 7.0 Kg Fully-Automatic Top Load Washing Machine", "Washing Machines", "Samsung", "WA70A4002GS", 19, 6, 13000, 16990, "Bay C-03", "3 Star", "3 Years", "Apex Electronics Distro", "In Stock", "2026-11-20 10:40"],
        ["XYZ-WM-204", "IFB 6.5 Kg 5-Star Fully-Automatic Front Load Washer", "Washing Machines", "IFB", "Diva Aqua SX", 0, 4, 21000, 27990, "Bay C-04", "5 Star", "4 Years", "Metro Appliance Wholesale", "Out of Stock", "2026-11-17 15:10"],
        ["XYZ-AC-301", "Daikin 1.5 Ton 5-Star Inverter Split Air Conditioner", "Air Conditioners", "Daikin", "FTKM50U", 16, 6, 36000, 45490, "Bay D-01", "5 Star", "1 Year / 10 Yrs Comp", "Pacific Tech Supplies", "In Stock", "2026-11-20 11:15"],
        ["XYZ-AC-302", "Voltas 1.4 Ton 3-Star Adjustable Inverter Split AC", "Air Conditioners", "Voltas", "173V Vectra Prism", 2, 8, 25000, 31990, "Bay D-02", "3 Star", "1 Year / 5 Yrs Comp", "Metro Appliance Wholesale", "Low Stock", "2026-11-20 12:25"],
        ["XYZ-AC-303", "Panasonic 1.5 Ton 5-Star Wi-Fi Inverter Smart Split AC", "Air Conditioners", "Panasonic", "CS/CU-NU18YKY5W", 9, 4, 34000, 43990, "Bay D-03", "5 Star", "1 Year / 10 Yrs Comp", "Apex Electronics Distro", "In Stock", "2026-11-20 09:50"],
        ["XYZ-KIT-401", "Philips Digital Air Fryer with Rapid Air Technology 4.1L", "Kitchen Appliances", "Philips", "HD9200/90", 25, 8, 5200, 7490, "Rack K-01", "N/A", "2 Years", "National Logistics Hub", "In Stock", "2026-11-20 14:00"],
        ["XYZ-KIT-402", "Bosch 13 Place Settings Free-Standing Smart Dishwasher", "Kitchen Appliances", "Bosch", "SMS66GI01I", 6, 4, 37000, 46990, "Bay E-01", "N/A", "2 Yrs / 10 Yrs Rust", "Pacific Tech Supplies", "In Stock", "2026-11-19 11:00"],
        ["XYZ-KIT-403", "Whirlpool 20L Solo Microwave Oven with Auto Cook Menu", "Kitchen Appliances", "Whirlpool", "MAGICOOKE20SE", 4, 10, 4500, 6290, "Rack K-02", "N/A", "1 Year", "Metro Appliance Wholesale", "Low Stock", "2026-11-20 13:45"],
        ["XYZ-KIT-404", "Prestige Iris Plus 750W Mixer Grinder with 4 Jars", "Kitchen Appliances", "Prestige", "Iris Plus", 35, 12, 2400, 3499, "Rack K-03", "N/A", "2 Years", "National Logistics Hub", "In Stock", "2026-11-20 15:10"],
        ["XYZ-KIT-405", "Morphy Richards 30L Convection Microwave Oven", "Kitchen Appliances", "Morphy Richards", "30 MCGR Deluxe", 0, 5, 11000, 14999, "Rack K-04", "N/A", "2 Years", "Apex Electronics Distro", "Out of Stock", "2026-11-18 17:30"],
        ["XYZ-CLN-501", "Dyson V12 Detect Slim Total Clean Cordless Vacuum", "Home Cleaning", "Dyson", "SV20", 7, 3, 44000, 55900, "Rack H-01", "N/A", "2 Years", "Pacific Tech Supplies", "In Stock", "2026-11-20 10:00"],
        ["XYZ-CLN-502", "Philips 1000i Series Smart Air Purifier HEPA Filter", "Home Cleaning", "Philips", "AC1215/20", 11, 5, 8500, 11999, "Rack H-02", "N/A", "2 Years", "National Logistics Hub", "In Stock", "2026-11-20 09:10"],
        ["XYZ-CLN-503", "Eureka Forbes Robo Vac n Mop Robotic Cleaner with LiDAR", "Home Cleaning", "Eureka Forbes", "LVI-100", 3, 5, 18000, 24990, "Rack H-03", "N/A", "1 Year", "Metro Appliance Wholesale", "Low Stock", "2026-11-19 14:15"]
    ]

    # Style definitions
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Navy Blue
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # Write headers
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border
    ws.row_dimensions[1].height = 28

    # Write data
    for row_idx, row_data in enumerate(data, 2):
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = data_font
            cell.border = thin_border
            
            # Alignments & formatting
            if col_idx in [1, 5, 10, 11, 12, 14, 15]:
                cell.alignment = center_align
            elif col_idx in [6, 7, 8, 9]:
                cell.alignment = right_align
                if col_idx in [8, 9]:
                    cell.number_format = '₹#,##0'
                else:
                    cell.number_format = '#,##0'
            else:
                cell.alignment = left_align

            # Color highlight for status
            if col_idx == 14: # Status
                if val == "In Stock":
                    cell.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid") # light green
                    cell.font = Font(name="Calibri", size=10, bold=True, color="166534")
                elif val == "Low Stock":
                    cell.fill = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid") # light yellow
                    cell.font = Font(name="Calibri", size=10, bold=True, color="854D0E")
                elif val == "Out of Stock":
                    cell.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid") # light red
                    cell.font = Font(name="Calibri", size=10, bold=True, color="991B1B")

        ws.row_dimensions[row_idx].height = 20

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    wb.save(file_path)
    print(f"Generated sample Excel at: {file_path}")

if __name__ == "__main__":
    target_path = r"D:\XYZ_Appliances_Stock_App\data\XYZ_Stock_Inventory.xlsx"
    create_sample_excel(target_path)
