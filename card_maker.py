# Script for drawing playing cards based on contents of an Excel file
# The file should contain three columns: level, front and back texts
#
# Wouter Bergmann Tiest
# 4 September 2024
#
# Requires openpyxl and pycairo:
#   pip3 install openpyxl
#   brew install cairo pkg-config
#   pip3 install pycairo

import openpyxl
import cairo
from math import pi

INPUT_FILE = 'words.xlsx' # data file with three columns
OUTPUT_FILE = 'cards.pdf'
PAGE_WIDTH = 420 * 72 / 25.4  # in points (1/72 inch); this is A3
PAGE_HEIGHT = 297 * 72 / 25.4
WIDTH = 240 # width of card, in points (1/72 inch)
HEIGHT = 160 # height of card, in points (1/72 inch)
RADIUS = 10 # corner radius, in points (1/72 inch)
INSET = 10 # margin between edge of card and graphic, in points (1/72 inch)
MARGIN = 15 # margin between page edge and cards, in points (1/72 inch)
CARDS_HORI = round(PAGE_WIDTH - 2 * INSET) // WIDTH # number of cards per page horizontally
CARDS_VERTI = round(PAGE_HEIGHT - 2 * INSET) // HEIGHT # number of cards per page vertically
CARDS_PER_PAGE = CARDS_HORI * CARDS_VERTI # total number of cards per page

# displays a text centered around (x, y)
def centered_text(ctx, x, y, text):
    ext = ctx.text_extents(text)
    ctx.move_to(x - ext.width / 2 - ext.x_bearing, y - ext.height / 2 - ext.y_bearing)
    ctx.show_text(text)

# splits a long text in parts narrower than max_width en displays it centered around (x, y)
def multi_line_text(ctx, x, y, text, max_width):
    split = text.split(' ')
    index = 0
    parts = []
    max_height = 0
    while index < len(split):
        part = split[index] # first word of next part
        ext = ctx.text_extents(part) # check size
        while ext.width <= max_width: # this loop is assumed to execute at least once
            last_part = part
            index += 1
            if index == len(split): # end of text reached?
                break
            part += ' ' + split[index] # try adding next word
            ext = ctx.text_extents(part) # check size again
        max_height = max(ext.height, max_height) # different lines might have different heights
        parts.append(last_part) # collect part without the last word added
    for i, part in enumerate(parts):
        centered_text(ctx, x, y + (i - len(parts) / 2 + 0.5) * max_height, part)

# creates the path for a rounded rectangle with top left corner (x, y)
def rounded_rectangle(ctx, x, y, width, height, radius):
    ctx.new_path()
    ctx.arc(x + radius, y + radius, radius, pi, 1.5 * pi) # top left
    ctx.arc(x + width - radius, y + radius, radius, 1.5 * pi, 2 * pi) # top right
    ctx.arc(x + width - radius, y + height - radius, radius, 0, 0.5 * pi) # bottom right
    ctx.arc(x + radius, y + height - radius, radius, 0.5 * pi, pi) # bottom left
    ctx.close_path()

# paints the red cutting line and the background in the given color
def background(ctx, color):
    rounded_rectangle(ctx, 0.1, 0.1, WIDTH - 0.2, HEIGHT - 0.2, RADIUS)
    ctx.set_source_rgb(1, 0, 0) # red
    ctx.set_line_width(0.01) # for laser cutter
    ctx.stroke()
    rounded_rectangle(ctx, INSET, INSET, WIDTH - 2 * INSET, HEIGHT - 2 * INSET, RADIUS)
    ctx.set_source_rgb(*color)
    ctx.fill_preserve()
    ctx.set_source_rgb(0, 0, 0) # black
    ctx.set_line_width(2)
    ctx.stroke()

# paint the front of the card
def make_front(surface, card):
    ctx = cairo.Context(surface)
    background(ctx, (0.9, 0.9, 1)) # light blue
    ctx.set_source_rgb(0.7, 0.7, 1) # somewhat darker blue
    ctx.select_font_face('sans', cairo.FONT_SLANT_NORMAL, cairo.FONT_WEIGHT_NORMAL)
    ctx.set_font_size(144)
    centered_text(ctx, WIDTH / 2, HEIGHT / 2, str(card['level']))
    ctx.set_source_rgb(0, 0, 0) # black
    ctx.set_font_size(24)
    multi_line_text(ctx, WIDTH / 2, HEIGHT / 2, card['term'], WIDTH - 3 * INSET)

# paints the back of the card
def make_back(surface, card):
    ctx = cairo.Context(surface)
    background(ctx, (0.9, 1, 0.9)) # light green
    ctx.set_source_rgb(0.7, 1, 0.7) # somewhat darker green
    ctx.select_font_face('sans', cairo.FONT_SLANT_NORMAL, cairo.FONT_WEIGHT_NORMAL)
    ctx.set_font_size(144)
    centered_text(ctx, WIDTH / 2, HEIGHT / 2, str(card['level']))
    ctx.set_source_rgb(0, 0, 0) # black
    ctx.set_font_size(12)
    multi_line_text(ctx, WIDTH / 2, HEIGHT / 2, card['definition'], WIDTH - 3 * INSET)

# displays a number in the bottom left corner of the card
def number(surface, i):
    ctx = cairo.Context(surface)
    ctx.set_source_rgb(0, 0, 0) # black
    ctx.set_font_size(6)
    ctx.move_to(1.5 * INSET, HEIGHT - 1.5 * INSET)
    ctx.show_text(str(i))

# load entire input file
cards = []
wb = openpyxl.load_workbook(INPUT_FILE)
ws = wb.active
for row in ws.iter_rows(min_row=2, values_only=True): # skip header row
    cards.append({'level': row[0], 'term': row[1], 'definition': row[2]})

# create output file
surface = cairo.PDFSurface(OUTPUT_FILE, PAGE_WIDTH, PAGE_HEIGHT)

# draw all cards
front = True # first draw a page of card fronts
i = 0
while i < len(cards):
    if front:
        # determine position, left to right, top to bottom
        x = MARGIN + (i % CARDS_HORI) * WIDTH
        y = MARGIN + ((i // CARDS_HORI) % CARDS_VERTI) * HEIGHT
        sub_surface = surface.create_for_rectangle(x, y, WIDTH, HEIGHT)
        make_front(sub_surface, cards[i])
        number(sub_surface, i + 1)
        if (i + 1) % CARDS_PER_PAGE == 0: # page full?
            surface.show_page()
            front = False # switch to drawing card backs
            i -= CARDS_PER_PAGE - 1 # reset counter to start of page
        elif i == len(cards) - 1: # last card when page is not yet full?
            surface.show_page()
            front = False # switch to drawing card backs
            i -= i % CARDS_PER_PAGE # reset counter to start of page
    if not front: # do not use 'else' here; both cases are executed at the turn of a page
        # determine position, right to left, top to bottom
        x = round(PAGE_WIDTH - MARGIN - (i % CARDS_HORI + 1) * WIDTH)
        y = MARGIN + ((i // CARDS_HORI) % CARDS_VERTI) * HEIGHT
        sub_surface = surface.create_for_rectangle(x, y, WIDTH, HEIGHT)
        make_back(sub_surface, cards[i])
        number(sub_surface, i + 1)
        if (i + 1) % CARDS_PER_PAGE == 0: # page full?
            surface.show_page()
            front = True # switch back to drawing card fronts
    i += 1
surface.finish()
