# 📦 VBA Macro – Service Level Calculation

This macro automates the calculation of the **logistics service level** from raw Excel data.

## 🔧 Features
- Reads a data table with delivery records
- Aggregates delivered and ordered volumes by **client** and **warehouse**
- Calculates the service level percentage
- Automatically generates a clean summary table in a new worksheet

## 📝 Expected Data Format
The source worksheet must be named `Données` and contain the following columns:

| A: Client | B: Warehouse | C: Order Date | D: Product | E: Ordered Quantity | F: Delivered Quantity (On Time) |

## 📁 Files Included
- `taux_service.bas` : the VBA source code
- `sample_data.xlsx` : a dummy Excel file to test the macro
- `README.md` : this file

## ▶️ How to Use It
1. Open `sample_data.xlsx`
2. Press `ALT + F11` to open the VBA editor
3. Go to `File > Import File...` and select `taux_service.bas`
4. Run the macro `CalculTauxService`
5. The results will be generated in a new sheet named `Résultat`

## 💡 Author
**Code by Ranim**  
✨ Follow for more practical automation tools using Excel + VBA!
