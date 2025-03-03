import os
import base64

# Create static directory if it doesn't exist
static_dir = os.path.dirname(__file__)
os.makedirs(static_dir, exist_ok=True)

# Simple base64-encoded icon for the add-in (blue square with "OR" text)
icon32_b64 = """
iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAABK0lEQVR42mNkGIqAEd3ntbW1/
9EwVqW1tbWMo3YMWQfU1dX9pyYeDAEYBQ7A5X9wayG7npKUAHfAUHIAC7q/8YU+LrigdmFEsQ
TOEBhFMOAAXFEQGRkJZmdnZ4MVbGxswl2Ym5sLVlRQUABm+/v7g9mXLl0apA5gRg//jo4OMBt
W8iEDWLTAQh4XGEDLUKKgoAAliYH4e/fuQZwwcA5AzgWwdI/NAfDEhy8aDAwMBq8DmJFzAXLa
R04DyLkAvRhGLsjHGwVoDsCXBoaj/wctCpBdAEv7sJIQW96AVT2jDqBuWTCKYPAyIbIjkNM5c
skHw0OpTzB0HcBAQNPAOmAUDYYoYCLFACMjI1w5rK06WN4IwQAWMBNGQ2BwRAG1AWPbkAIAAL
wF27iJXhCHAAAAAElFTkSuQmCC
"""

icon80_b64 = """
iVBORw0KGgoAAAANSUhEUgAAAFAAAABQCAYAAACOEfKtAAAABHNCSVQICAgIfAhkiAAAABl0RVh0
U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZy8yMDIxLz+/LSwAAATTSURBVHic7ZxNiBxFFMd/1T0z
u9mNJEYTsol4MIKIaLwIXgQPQTx4ULyIePAgeBVPgnhVEFQ8iQcP4smDF/EieBBERBBEVEQSxWiS
TYxZs53dmZ7uejxMTaZ3Zz7eTE1Xb3f1D3roTPrV65r3b1e9V6+qhYiwlZFsdAe6xlZ3IBvdAYCi
KPDbNE0pigJf94Wi+NdtEW0AiAhJkqz5m/87jmMkScJGOnLDCez3+3HnXuwC58M30P+HMTIqlcqG
9eeGCCzLkjRNAej1emRZRqlCbVxKEiUKILEKKveE3IC6rlAPePrCrr4p8QpjK5Zmf+CNz77lxuqa
FzALfs+8/AzXT87QE5VeMC4BLvwtC/WXbSAxtWMq6LE03xWBuYiUohJ3JTwGDc/oVDm7+z6OQ1wC
pzp2AJyFPxaRrv8atXtWwTgE2kox7ApuNYxIYNZv8MnprzlxdqbKgwxPRapRPqIYDdmLh6pacvG5
47WWwK1i4GgFal3RuGUHRBi8riEigcvlY/8Vtgwaz2SvVyedOkpT5HrrLzBLFPQonGf58qZgAlXy
+OPPRtueoJGf6qPpr9HPRqpGi6re+iXQkldFRaNnHwVRGAdTYB69AKkv+lqDCNPGS9AYsEIgotx2
w4Tt/ihp0iNNFO99+CkLi0urKiyZQADVwueOCtsbFtMWepGYImsJJLCJy8CoGjxw3/5oi1NrCOxq
5SwH9kewwP4b4JnDz8Un0KavxkprCyrTf3lt5/jYznGWllaXgGkqGjOV/fjEifbPQgOjqKSbd1Gv
kF0S118RSRvF00pzIa3wFaatQw2c+uoktd9Y+3OEaaufQjX9ooqGRXcEDgVpuUGpA9n3xQp8MemL
xShGQbLtQJYZ5hbP89aRo8xdrC4QDX4cIMvsj39++/0ME2M7eO7Zw1ztrK9htRsvw2praWEJfvnp
DP9PRiSQprqHH9jHwUdur9ZSqciOzP50hlOzP1BXRkTr05uBrhF6jbr5lr2MZBnYHm7vrffnWVhc
rLwgo/P0cwARWLzQYWHxUvt7UASPwi48f+QYl5dXllnHQUtvBPbLjvfefJnqqsxg/3A257UogZ6N
9T5pbEZzQ2SDsf6vIN1HZ5ckQu+NX9HMQ2BnF39DgkrGGobYvcEFleWvbcZNpwgzsO3iqnG1xyPB
JHeXZjaUQaAxmhx7+729fO7BXXMk9hc32+rW0zZVaRjlioRA02+YO7uAxMp6bDIrpQVLtqHWALQj
cTavJ3AiDJuWkteRl6EpBMsiUjua08DkLK2sKlbNiV3LRus1TiN+h7AYneaNNjf55dffee+jY2QD
JBNaQyPMZJqO07QOTWvBr779CYDFpRVmc89LDokemj8/Is3O5NS+Ru9n46bE/rlwGxRirgZumTnT
PoHrOYNGC+fURJSGKxT+D2ucisYCdqVymu2eMjuxgx9/+qn5rzNr7outS11q83qgwOXV5lwR8Z5L
/oi5FmiiP8M5NgZuhey1Heio+e7EhkXHOz9MVBo3JlTaJJ5AGDu9yzn4MiLFFUe9hd2+NS77IBZG
EElkDzvK9S3yvMJZwx3Lp4W/AhP9+WA9AuMvdcmo9o1DTwsjEViaJZx1r+q+3IpOq7pbKK0pT8v6
7h6cw8aE/KtPD+KcTfyLnwwveRKprr65WA9xhz9faf5xI6wH7sFvCXXKImht8bVREeXldxty2+Fw
jOHoBUQX8IuXKN2hLBu+OIFwwV5gbdNL5NjPLYcbzGqjKdZCJG//+V7Tfs/QFwAAAABJRU5ErkJg
gg==
"""

# Create icon files if they don't exist
icon32_path = os.path.join(static_dir, "icon-32.png")
icon80_path = os.path.join(static_dir, "icon-80.png")

try:
    if not os.path.exists(icon32_path):
        with open(icon32_path, 'wb') as f:
            f.write(base64.b64decode(icon32_b64))
        print(f"Created {icon32_path}")
    
    if not os.path.exists(icon80_path):
        with open(icon80_path, 'wb') as f:
            f.write(base64.b64decode(icon80_b64))
        print(f"Created {icon80_path}")
    
    print("Icon files created successfully!")
except Exception as e:
    print(f"Warning: Could not create icon files: {str(e)}")
