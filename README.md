# Blackjack Data 

## Dependencies
Python
```shell script
pip install xlsxwriter tqdm xlrd
```

## How to use

Clone de repo
``` shell
git clone https://github.com/obin1000/Blackjack_Data_Research
```
Enter the new folder
``` shell
cd Blackjack_Data_Research/
```
Generate the data set with python, providing the size of the set (in this example 100000)
``` shell
python blackjack_data_generator.py 100000
```
Refine the generated data set with python
``` shell
python blackjack_data_refiner.py
```
