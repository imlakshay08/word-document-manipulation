# Word File Data Manipulation and Export Project

This project, I developed for [ZeroAndOne Developers](https://zodevelopers.com/) community, is aimed at efficiently manipulating specific data within a Word document and exporting it as a PDF. As a Software Development Associate at ZeroAndOne Developers, I've been tasked with completing this project to meet the requirements of our clients. The application is designed to replace text while maintaining formatting, perform calculations within table columns, and integrate a user-friendly Streamlit dashboard for input handling.

## Features

- **Text Replacement**: Replace text within paragraphs and tables while preserving formatting.
- **Calculation in Tables**: Perform calculations within table columns based on user-defined factors.
- **PDF Export**: Export the modified Word document as a PDF.
- **Streamlit Dashboard**: Integrate a Streamlit dashboard for user input and interaction.

## Getting Started

To use this application, follow these steps:

1. Clone the repository to your local machine.

```bash
   git clone <repository_url>
   ```
2. Install the required dependencies. Ensure you have Python installed on your system.

```bash
pip install -r requirements.txt
```
3. Run the application using the following command:

```bash
streamlit run main.py
```
4. Enter the password when prompted. The default password is stored securely.

5. Fill in the required details in the Streamlit dashboard:
  - Company Name
  - Address
  - Factor (for calculations)
6. Click on the "Generate Documents" button to generate the modified Word document and export it as a PDF.

## Technical Details
- **Python Libraries**:
 - streamlit: Used for creating the user-friendly dashboard.
 - python-docx: Utilized for manipulating Word documents.
 - comtypes: Employed for PDF export functionality.

- **Code Organization**:
 - main.py: Contains the main application code, including Streamlit dashboard setup, document manipulation, and PDF export functionality.
 - Template_for_word_replace.docx: Template Word document for data replacement.
 - requirements.txt: List of Python dependencies.

## Usage
1. **Password Protection**:
 - Enter the password when prompted to access the application.
2. **Streamlit Dashboard**:
 - Input Company Name, Address, and Factor in the provided fields.
 - Click on the "Generate Documents" button to initiate the document   modification process.
3. **Document Modification**:
 - The application replaces placeholders in the document with user-provided data.
 - It performs calculations within table columns based on the specified factor.
4. **PDF Export**:
 - The modified document is saved as modified_document.docx.
 - It is then converted to PDF format and saved as final.pdf.

## Support

For any issues or inquiries, please open an issue in the GitHub repository or contact .

## Contribution

Contributions to the project are welcome! Feel free to open pull requests with improvements or additional features.

## License
This project is licensed under the MIT License.
