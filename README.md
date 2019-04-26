# receipt-extractor

`fetch.rb`: Downloads receipts from IMAP servers and intelligently saves them into PDF files.

`extract.rb`: Parses PDF files and extracts supplier, invoice date and amount into a tab-separated output.

## Configuration

Edit `fetch.rb` to add your IMAP server data and the appropriate filters and extractors for your use case. Edit `extract.rb` to add the additional parsers for the generated PDFs as needed.

## Usage

Run `fetch.rb`. It will save PDF files to the current working directory, named with timestamp, sender and message ID.

Then, run `extract.rb` passing the PDF files as args and redirecting the output to a `.tsv` file.

Open the `.tsv` file with Excel and copy-paste to your heart’s content.
