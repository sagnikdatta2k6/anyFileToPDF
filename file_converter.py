import os
from fpdf import FPDF
from PIL import Image
from docx import Document
import openpyxl

try:
    import comtypes.client
    POWERPOINT_INSTALLED = True
except ImportError:
    POWERPOINT_INSTALLED = False
