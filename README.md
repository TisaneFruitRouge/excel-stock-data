# Excel Stock Data



This program was made for the user **petee0328** on [Fiverr.com](https://www.fiverr.com/petee0328)

There are 3 main functions: 

- *main()* which goes thought the data in the excel file
- *create_transaction_json(shares_dict)* which, given a dictionary creates a JSON file for each share, bought and sold, with the corresponding amount, price and profit made.
- *create_balance_json(profit_dit)* which, given a dictionary containing the list of profit for each share for a given transaction, updates the balane.json file.
- *get_round_value(float_number)* which rounds a float number to 2 digits after the decimal point. (didn't use the built-in *round()* function because its behaviour is sometimes wrong).

Prices and profits (in dollars) are rounded to 2 digits after the decimal point.