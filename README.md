# Word Document Processor

This Spring Boot application allows you to:
1. Upload a Word document (.docx format)
2. Prefill placeholders with specified values
3. Add additional rows to tables if needed
4. Export the modified document

## Prerequisites
- Java 11 or higher
- Maven

## Running the Application

1. Clone the repository
2. Navigate to the project directory
3. Run the application using Maven:
   ```
   mvn spring-boot:run
   ```
4. Open your browser and go to `http://localhost:8080`

## How to Use

1. Upload your Word document (.docx format)
2. Specify the number of additional rows you want to add to the table (if any)
3. Add placeholder-value pairs:
   - In the "Placeholder text" field, enter the text you want to replace
   - In the "Replacement value" field, enter the new value
   - Click "Add Another Replacement" to add more pairs
4. Click "Process Document" to generate the modified document
5. The processed document will be automatically downloaded

## Notes

- The application supports multiple replacements in the same document
- Make sure your Word document is not password-protected
- The application preserves the original formatting of the document
- Additional rows will inherit the formatting of the last row in the table
