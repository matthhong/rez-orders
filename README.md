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

d) The master.csv file will have been created. In Excel (or whatever you're comfortable with), order the rows of master.csv in the way you want them to be eventually ordered in the stickers file. While you are at it, skim through all rows to make sure there are no glaring errors. 

e) Along with the master.csv, data with known errors were saved in bad_orders.csv. You will need to manually add the information in this file into master.csv. Sometimes you will have to email those people to correct the errors (usually missing information).

f) Run:
```
python order-exports.py stickers
```

g) The stickers will have been created in labels.docx. It is your responsibility to check that they do not contain errors.
