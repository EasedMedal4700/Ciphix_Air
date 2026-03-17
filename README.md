# Ciphix RPA Case – Letter of Notice Automation

## 📌 Overview
This project is an RPA solution built in UiPath for Ciphix Air.  
The goal of the automation is to generate and send a **Letter of Notice (LoN)** to passengers whose lost items have been found.

A key challenge in the process is **missing postal codes**, which prevents successful delivery of the letter.  
This automation resolves that issue by retrieving missing postal codes using the PostNL postcode lookup service.

---

## 🎯 Objectives
- Read and clean input data from Excel
- Validate passenger and address data
- Retrieve missing postal codes via PostNL
- Generate a Letter of Notice using a Word template
- Convert the document to PDF
- Log invalid records and handle exceptions

---

## ⚙️ Process Flow

1. **Read Input Data**
   - Load `InputTransactionData.xlsx`

2. **Data Cleaning**
   - Trim spaces
   - Normalize casing
   - Remove unwanted characters

3. **Validation**
   - Validate required fields (name, address, flight data)
   - Separate invalid records

4. **Postal Code Retrieval**
   - Use PostNL website:
     https://www.postnl.nl/postcode-zoeken/
   - Input: StreetName, HouseNumber, CityName

5. **Generate Letter of Notice**
   - Fill `letterofnotice.docx` using bookmarks

6. **Convert to PDF**
   - Save as:
     ```
     LetterOfNotice_{LastName}_{FirstName}_{PassengerId}_{FlightNo}.pdf
     ```

7. **Logging & Exception Handling**
   - Log invalid rows
   - Continue processing remaining records

---

## 🧱 Architecture
The automation follows a **transaction-based design**:
- Each row is treated as an individual transaction
- Invalid rows are skipped and logged
- Processing continues independently per row

---

## 📁 Project Structure
