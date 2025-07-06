import re
import pandas as pd
import pytesseract
from PIL import Image


def conv_im_to_ex(img):


    img = Image.open(img)

    text = pytesseract.image_to_string(img)


    # Split the text into lines
    lines = text.split('\n')

    # Initialize lists to store extracted data
    data = []
    for line in lines:
        if re.match(r"\d{2}/\d{2}", line):  # Check if the line starts with a date-like pattern
            parts = line.split()
            if len(parts) >= 3:
                date = parts[0]
                operation = " ".join(parts[1:-2])
                amount = parts[-1].replace(",", ".")
                if "Débit" in operation:
                    debit = amount
                    credit = "0.00"
                else:
                    debit = "0.00"
                    credit = amount
                data.append([date, operation, debit, credit])

    # Create a DataFrame from the extracted data
    columns = ["Date", "Opérations", "Débit", "Credit"]
    new_df = pd.DataFrame(data, columns=columns)

    # Read the existing Excel file into a DataFrame (if it exists)
    try:
        existing_df = pd.read_excel("dataset.xlsx")
    except FileNotFoundError:
        existing_df = pd.DataFrame(columns=columns)

    # Concatenate the new data with the existing DataFrame
    combined_df = pd.concat([existing_df, new_df], ignore_index=True)

    # Save the combined DataFrame to the Excel file
    combined_df.to_excel("dataset.xlsx", index=False)
    print("Data added to Excel file successfully.")
