# ShaiMauiExcelToSql

A robust .NET MAUI Blazor Hybrid application designed to bridge the gap between SQL Server databases and Excel reporting. This tool allows users to execute SQL queries directly against a database and export the results into formatted, professional-looking Excel spreadsheets.

## 🚀 Features

## 📥 Download
Latest Release: [Download for Windows (x64)](https://github.com/shadijhade/ShaiMauiExcelToSql/releases/latest)


-   **SQL Connectivity**: Connect to any SQL Server database using custom connection strings.
-   **Query Execution**: Run standard SQL queries and view results immediately within the application.
-   **Advanced Excel Export**:
    -   Export single or multiple result sets to Excel.
    -   **Automatic Formatting**: Applies table styles, auto-fits columns, and freezes headers for better readability.
    -   **Customization**: Users can choose from various Excel table styles (Light, Medium, Dark).
    -   **Data Integrity**: Handles various data types and DBNull values gracefully.
-   **Multi-Result Set Support**: capable of handling and exporting complex queries that return multiple tables.
-   **Query History**: Keeps track of executed queries for quick re-use (in-memory/session based).
-   **Cross-Platform**: Built on .NET MAUI, capable of running on Windows (focused), macOS, iOS, and Android.

## 🛠️ Technology Stack

-   **Framework**: [.NET MAUI](https://dotnet.microsoft.com/en-us/apps/maui) with **Blazor Hybrid**
-   **Language**: C# / Razor
-   **Database Client**: `Microsoft.Data.SqlClient`
-   **Excel Generation**: [ClosedXML](https://github.com/ClosedXML/ClosedXML)
-   **UI/UX**: HTML/CSS/Razor components powered by `Microsoft.AspNetCore.Components.WebView.Maui`
-   **Utilities**: `CommunityToolkit.Maui`

## 📋 Prerequisites

Before you begin, ensure you have the following installed:

-   [Visual Studio 2022](https://visualstudio.microsoft.com/vs/) (17.8 or later)
-   **.NET Multi-platform App UI development** workload installed in VS.
-   [.NET 8 SDK](https://dotnet.microsoft.com/en-us/download/dotnet/8.0)
-   A SQL Server instance (LocalDB, Express, or standard) to connect to.

## 🏃 Getting Started

1.  **Clone the Repository**
    ```bash
    git clone https://github.com/shadijhade/ShaiMauiExcelToSql.git
    cd ShaiMauiExcelToSql
    ```

2.  **Open the Solution**
    -   Launch Visual Studio 2022.
    -   Open `ShaiMauiExcelToSql.sln`.

3.  **Restore Dependencies**
    -   Right-click on the solution in Solution Explorer and select **Restore NuGet Packages**.

4.  **Run the Application**
    -   Select **Windows Machine** as the target device.
    -   Press `F5` or click **Run**.

## 📖 Usage Guide

1.  **Connect**: Enter your SQL Server connection string in the configuration area.
2.  **Query**: Type your SQL query (e.g., `SELECT * FROM Users`) in the query editor.
3.  **Execute**: Click the **Run** button to fetch data.
4.  **View Results**: Data will appear in the grid view. If multiple result sets are returned, they will be displayed in tabs or sections.
5.  **Export**: Click the **Export to Excel** button.
    -   Select your desired table style.
    -   Choose the destination folder.
    -   The app will generate an `.xlsx` file with your data, formatted and ready to use.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

This project is open-source and available under the [MIT License](LICENSE).