# rez-orders
Instructions:

a) In the command line, type:
```
pip install python-docx
```

b) Place the 'orders_export.csv' file in the same folder as orders-export.py and landscape.docx.

c)
```
python order-exports.py master
```

d) The master.csv file will have been created. Order the rows in the way you want them to be eventually ordered in the stickers file. While you are at it, skim through all rows to make sure there are no errors. 

e) Rows where there are known errors are saved in bad_orders.csv. You will need to manually add the information in this file into master.csv.

f)
```
python order-exports.py stickers
```

g) The stickers will have been created in labels.docx. It is your responsibility to check that they do not contain errors.
