import sys

# Completely disable dummy/incompatible numpy and pandas in frozen bundle
# This forces openpyxl and reportlab to use pure Python numeric types (int, float, Decimal)
# and prevents any dummy/stub numpy module from raising AttributeError on missing types (intc, short, etc.)
sys.modules['numpy'] = None
sys.modules['pandas'] = None
