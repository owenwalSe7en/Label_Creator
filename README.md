# ERP Automation Tool

Welcome to the ERP Automation Tool! This project is designed to automate the process of creating profiles for large amounts of data by reading an external Excel file and writing the part numbers into the UI of a specific ERP system. The tool leverages `pywinauto` for interacting with the ERP system's UI, `openpyxl` for reading and writing Excel files, and `tkinter` for providing a user-friendly interface.

## Features

- **Automates Profile Creation**: Automatically creates profiles in the ERP system using data from an Excel file.
- **Excel Integration**: Reads part numbers and other necessary data from an external Excel file.
- **User-Friendly UI**: Uses `tkinter` to provide a simple and intuitive interface for connecting the Excel file and configuring the program.
- **Error Handling**: Includes robust error handling to manage common issues during automation.

## Installation

1. **Clone the repository**
    ```bash
    git clone https://github.com/owenwalSe7en/LabelCreator.git
    cd LabelCreator
    ```

2. **Create a virtual environment**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

3. **Install dependencies**
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1. **Run the application**
    ```bash
    python main.py
    ```

2. **Configure the tool via the UI**
    - **Select Excel File**: Use the provided UI to navigate and select the Excel file containing the part numbers.
    - **Specify Parameters**: Enter the necessary details for the program to connect to the Excel file and the ERP system.

3. **Start Automation**
    - Click the "Submit" button in the UI to begin the automation process. The tool will read the part numbers from the selected Excel file and input them into the ERP system.

## Dependencies

- **pywinauto**: For interacting with the ERP system's UI.
- **openpyxl**: For reading from and writing to Excel files.
- **tkinter**: For creating the graphical user interface.

## File Structure

- `main.py`: The main script to run the application.
- `LC_Automation.py`: Contains the code for gathering user-inputed data and automating the profile creation in the ERP system.
- `requirements.txt`: Lists the Python dependencies required for the project.

## Contributing

Contributions are welcome! Please follow these steps to contribute:

1. **Fork the repository**
2. **Create a new branch** for your feature or bugfix
    ```bash
    git checkout -b feature/your-feature-name
    ```
3. **Commit your changes**
    ```bash
    git commit -m 'Add some feature'
    ```
4. **Push to the branch**
    ```bash
    git push origin feature/your-feature-name
    ```
5. **Create a new Pull Request**

## Contact

For any questions or suggestions, please feel free to open an issue or contact me at [wallaceowenh45@gmail.com](mailto:your-email@example.com).
