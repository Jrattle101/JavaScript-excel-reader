# Excel to SharePoint App

This JavaScript application allows you to upload an Excel file, read its data, and add it to a SharePoint list. It also checks for duplicates in the SharePoint list before adding the data.

## Features

- Upload Excel files.
- Read Excel data and convert it to JSON format.
- Add data to a SharePoint list.
- Check for duplicates in the SharePoint list and scrub them out.

## Technologies Used

- JavaScript
- SheetJS (for reading Excel files)
- SharePoint REST API

## Usage

1. Clone the repository:

    ```bash
    git clone https://github.com/Jrattle101/JavaScript-excel-reader.git
    ```

2. Open `main.html` in your preferred web browser.

3. Click on the "Choose File" button and select an Excel file.

4. Click on the "Upload to SharePoint" button to upload the data to your SharePoint list.

## Configuration

- Ensure you have the necessary permissions to access SharePoint lists and use the SharePoint REST API.
- Modify the SharePoint site URL and list name in the JavaScript code to match your SharePoint environment:

    ```javascript
    var siteUrl = "https://your-sharepoint-site-url";
    var listName = "YourListName";
    ```

## Contributing

Contributions are welcome! If you have any suggestions, bug reports, or feature requests, please open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).
