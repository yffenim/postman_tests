# Turn a Postman Collection into a table of request results*

*For when one does not have API testing set up..

## Usage

Ensure you have python3 installed.

Run the install: `pip3 install requirements.txt`

Put your exported Postman JSON file into `/collections_json/`

Run the script: `python3 validate_endpoints.py`

When its done, you will have both a log of the results and a word doc w/ separate tables per collection, the 404s and Errors are bolded, and there's columns you can rename.

