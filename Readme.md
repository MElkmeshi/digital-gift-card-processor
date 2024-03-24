# Digital Gift Card Processor

The Digital Gift Card Processor automates the conversion of digital gift card data from text files into a unified Excel format. This tool is designed to streamline the process of preparing gift card data for server uploads, especially useful for businesses frequently dealing with gift card batches from providers like Almadar, Libyana, or LTT. By converting text files into a specific Excel format accepted by servers, this utility saves significant manual processing time.

## Features

- **Batch Processing**: Automatically processes multiple gift card text files in one go, organizing them into specified Excel formats.
- **Custom Value Mapping**: Maps specific gift card values to corresponding categories, ensuring accurate data representation.
- **Dynamic Excel Creation**: Generates Excel files ready for server upload, significantly reducing manual data entry and potential errors.
- **Flexible Directory Support**: Works with gift card data spread across different folders, accommodating varied organizational structures.

## Getting Started

### Prerequisites

- Rust Programming Environment: Ensure you have Rust installed on your system. If not, follow the [official Rust installation guide](https://www.rust-lang.org/tools/install).
- `xlsxwriter` Crate: For generating Excel files.
- `csv` Crate: For reading CSV files.

### Installation

1. Clone the repository to your local machine:

   ```
   git clone https://github.com/melkmeshi/digital-gift-card-processor.git
   ```

2. Navigate to the project directory:

   ```
   cd digital-gift-card-processor
   ```

3. Build the project:

   ```
   cargo build --release
   ```

4. The executable will be available under `./target/release/`.

### Usage

1. Prepare your text files containing the digital gift card data in the specified format.
2. Run the processor:

   ```
   ./target/release/digital_gift_card_processor
   ```

3. The program will automatically detect text files in the current directory (or specified directory), process them, and generate Excel files in the designated output location.

Certainly! Here's how you can update the README with a section that provides a clear example demonstrating the conversion from text data to the Excel format.

---

# Examples

## From Text Data to Excel

The Telecom SIM Card Data Processor transforms text files containing SIM card data into a formatted Excel spreadsheet. Below is an example demonstrating this conversion.

### Text Data from Libyana (5.000LYD__2020-03-12_.out.dec)

```
309086936642254,8948042436708,5.000 LYD,20311
309086936642254,8948042436708,5.000 LYD,20311
309086936642254,8948042436708,5.000 LYD,20311
```

This text represents a list of Gift cards, where each line contains the card's secret number, sequence number, value, and an additional code, all separated by commas.

### Converted Excel (Libyana 5.xlsx)
| CARD_SEQ       | CARD_SECRET      | Value_Card | Code_Card | Exp_Card | Com_Name |
|----------------|------------------|------------|-----------|----------|----------|
| 8948042436708  | 309086936642254  | 5          |           |          |          |
| 8948042436708  | 309086936642254  | 5          |           |          |          |
| 8948042436708  | 309086936642254  | 5          |           |          |          |

The processor reads the text data and splits each line by commas. It then writes the data into corresponding columns within an Excel file, as shown above.

The `Value_Card` field is extracted from the third column by trimming off the currency notation "LYD" and converting the value to a number. As the `Code_Card`, `Exp_Card`, and `Com_Name` fields are not provided in the input, they are left blank in the output.

To perform this conversion, run the processor in the directory containing your text files:

```
digital-gift-card-processor
```

After processing, the Excel file (`Libyana 5.xlsx`) is generated and ready for use.

---

This example in the README helps users to understand the functionality of your tool and sets expectations for the input and output of the process. Adjust the paths and filenames according to your actual usage.

## Another Example

Suppose you have received several text files from Almadar, each containing gift card data. Each text file's name corresponds to a gift card value and resides in its designated folder. Here's how the processor handles these files:

1. **File Organization**:

   ```
   ./gift_cards/Almadar/20200012030046_46745_105.csv
   ./gift_cards/Libyana/5.000LYD__2020-03-12_.out.dec
   ./gift_cards/LTT/m.kmeshi_mkcard_U202002101052744004_23243543.txt
   ```

2. **Running the Processor**:

   ```
   digital_gift_card_processor
   ```

3. **Output**:

   After processing, for each text file, an Excel file is created in the same directory:

   ```
   ./gift_cards/Almadar/Almadar 5.xlsx
   ./gift_cards/Libyana/Libyana 5.xlsx
   ./gift_cards/LTT/LTT 5.xlsx
   ```

   Each Excel file contains a sheet with the processed gift card data, ready for server upload.

## Contributing

Contributions are what make the open-source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

Distributed under the MIT License. See `LICENSE` for more information.

## Contact

Mohamed Elkmeshi - [@melkmeshi](https://twitter.com/melkmeshi) - elkmeshi2002@gmail.com

Project Link: [https://github.com/melkmeshi/digital-gift-card-processor](https://github.com/your-username/digital-gift-card-processor)