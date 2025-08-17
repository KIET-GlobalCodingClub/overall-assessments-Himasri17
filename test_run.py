import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import textwrap

# Test basic functionality
print("Testing application components...")

# Test 1: Check if Excel file can be read
try:
    df = pd.read_excel("students.xlsx")
    print("✓ Excel file loaded successfully")
    print("  Columns:", list(df.columns))
    print("  Rows:", len(df))
except Exception as e:
    print("✗ Error loading Excel:", e)

# Test 2: Check if template exists
if os.path.exists("certificate_template.png"):
    print("✓ Certificate template found")
else:
    print("✗ Certificate template not found")

# Test 3: Check if font exists
if os.path.exists("fonts/Arial.ttf"):
    print("✓ Font file found")
else:
    print("✗ Font file not found")

# Test 4: Check output directory
if os.path.exists("output"):
    print("✓ Output directory exists")
else:
    print("✗ Output directory not found")

# Test 5: Try to generate one certificate
try:
    template = Image.open("certificate_template.png")
    print("✓ Template image loaded successfully")
    print("  Image size:", template.size)
except Exception as e:
    print("✗ Error loading template:", e)

print("\nTest completed!")