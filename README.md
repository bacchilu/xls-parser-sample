This is a really simple parsing of an XLS file using the xlrd library.

The interesting part is how I implemented a "Visitor Design Pattern" to export in JSON format or in the PHP serialization (serialize/unserialize) format.

"generator.py" is used to randomly generate an Excel file for my usage.