# Automating Alternates on Digchip.com
Automating cross reference search on digchip.com

## Usage
1) Download Micrsoft Edge webdriver from from https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
<br>Note: Make sure Web Driver version matches exact build type of Edge

2) Install all dependencies. Selenium is the most important. <br> ```pip install selenium```

3) Replace file_path with file with all part numbers

4) Replace row name with row name with all part number

5) open cmd, go the directory and run using <br>```python separate_alt.py```


## Warnings
Code is slow by design to avoid timeouts from digchip. Each query takes 25-30 seconds
