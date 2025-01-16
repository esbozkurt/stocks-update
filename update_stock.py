import pandas as pd

# Excel dosyasını yükle
file_path = "Stock_Cargo_History.xlsx"
excel_data = pd.ExcelFile(file_path)

# Sayfaları yükle
stock_df = excel_data.parse("Stock")
cargo_transactions_df = excel_data.parse("Cargo Transactions")
history_df = excel_data.parse("History")

# İşlemleri uygulama
for index, row in cargo_transactions_df.iterrows():
    if pd.notna(row["Brand"]) and pd.notna(row["Model"]) and pd.notna(row["Quantity Taken"]):
        brand = row["Brand"]
        model = row["Model"]
        quantity_taken = row["Quantity Taken"]
        
        # Stok güncellemesi
        stock_index = stock_df[(stock_df["Brand"] == brand) & (stock_df["Model"] == model)].index
        if not stock_index.empty:
            stock_df.loc[stock_index, "Stock Quantity"] -= quantity_taken

        # Geçmişe ekleme
        history_entry = {
            "Date": row["Date"],
            "Action": "Cargo Transaction",
            "Details": f"{quantity_taken} units of {brand} - {model} taken by {row['Cargo Vehicle']}"
        }
        history_df = pd.concat([history_df, pd.DataFrame([history_entry])], ignore_index=True)

# Güncellenmiş dosyayı kaydet
output_path = "Updated_Stock_Cargo_History.xlsx"
with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
    stock_df.to_excel(writer, sheet_name="Stock", index=False)
    cargo_transactions_df.to_excel(writer, sheet_name="Cargo Transactions", index=False)
    history_df.to_excel(writer, sheet_name="History", index=False)

print(f"Updated Excel file saved as {output_path}")
