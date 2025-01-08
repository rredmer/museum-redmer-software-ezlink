# RSC EZ-LINK

RSC EZ-LINK is a downloader application for Symbol PDT Scanners developed by Redmer Software Company. This application facilitates the transfer and management of data between Symbol PDT Scanners and a central database.

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Contributing](#contributing)
- [License](#license)

## Features

- Download data from Symbol PDT Scanners using XMODEM protocol.
- Manage and store data in a Microsoft Access database.
- View and edit downloaded data.
- Serial communication with configurable ports.
- User-friendly interface for data management.

## Installation

### Prerequisites

- Microsoft Windows
- Microsoft Visual Basic 6.0
- Microsoft Access
- Symbol PDT Scanner

<img src="Images/Symbol PDT Barcode Scanning Point of Sale Terminal.gif" width="400" />

### Steps

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/museum-redmer-software-ezlink.git
    ```
2. Open the project in Microsoft Visual Basic 6.0.
3. Build the project to generate the executable.

## Usage

1. Connect the Symbol PDT Scanner to your computer.
2. Run the `EZLINK.exe` application.
3. Configure the serial port settings as required.
4. Use the interface to download and manage data from the scanner.

## Configuration

The application uses a Microsoft Access database (`ezlink.mdb`) to store configuration and log data. The database includes the following tables:

- `tbl_Configuration`: Stores application configuration settings.
- `tbl_Log_Info`: Stores log information of downloaded data.
- `tbl_Serial_Ports`: Stores available serial ports.

### Global Constants

The following global constants are defined in the application:

- `EZ_CAPTION`: Standard message box window caption.
- `EZ_APPDATA`: Name of the application database.
- `EZ_APPCONFIG`: Name of the configuration table.
- `EZ_APPLOG`: Name of the log info table.
- `EZ_APPPORTS`: Name of the serial port list table.
- `EZ_MSG_TECH_SUPPORT`: Technical support message.
- `EZ_CONSTRING`: Connection string for the database.

## Contributing

Contributions are welcome! Please follow these steps to contribute:

1. Fork the repository.
2. Create a new branch:
    ```sh
    git checkout -b feature-branch
    ```
3. Make your changes and commit them:
    ```sh
    git commit -m "Add new feature"
    ```
4. Push to the branch:
    ```sh
    git push origin feature-branch
    ```
5. Open a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.