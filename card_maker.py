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

# change these to customize
INPUT_FILE = 'words.xlsx' # data file with three columns
OUTPUT_FILE = 'cards.pdf'
SEPARATE_CUTTING_LINES = True # create a seprate file with only cutting lines for a laser cutter
CUTTING_FILE = 'cutting_lines.pdf'
PAGE_WIDTH = 420 * 72 / 25.4 # in points (1/72 inch); this is A3
PAGE_HEIGHT = 297 * 72 / 25.4
WIDTH = 240 # width of card, in points (1/72 inch)
HEIGHT = 160 # height of card, in points (1/72 inch)
RADIUS = 10 # corner radius, in points (1/72 inch)
INSET = 10 # margin between edge of card and graphic, in points (1/72 inch)
MARGIN = 15 # margin between page edge and cards, in points (1/72 inch)
DOTS = True # draw alignment dots
DOT_RADIUS = 8 if DOTS else 0 # radius of alignment dots

# do not change values below
NO = 0
YES = 1
ONLY = 2
CARDS_HORI = round(PAGE_WIDTH - 2 * MARGIN - 4 * DOT_RADIUS) // WIDTH # number of cards per page horizontally
CARDS_VERTI = round(PAGE_HEIGHT - 2 * MARGIN) // HEIGHT # number of cards per page vertically
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

# paints the red cutting line
def cutting_line(ctx):
    rounded_rectangle(ctx, 0.01, 0.01, WIDTH - 0.02, HEIGHT - 0.02, RADIUS)
    ctx.set_source_rgb(1, 0, 0) # red
    ctx.set_line_width(0.01) # for laser cutter
    ctx.stroke()

# paints the background in the given color
def background(ctx, color):
    rounded_rectangle(ctx, INSET, INSET, WIDTH - 2 * INSET, HEIGHT - 2 * INSET, RADIUS)
    ctx.set_source_rgb(*color)
    ctx.fill_preserve()
    ctx.set_source_rgb(0, 0, 0) # black
    ctx.set_line_width(2)
    ctx.stroke()

# paint the front of the card
def make_front(ctx, card):
    background(ctx, (0.9, 0.9, 1)) # light blue
    ctx.set_source_rgb(0.7, 0.7, 1) # somewhat darker blue
    ctx.select_font_face('sans', cairo.FONT_SLANT_NORMAL, cairo.FONT_WEIGHT_NORMAL)
    ctx.set_font_size(144)
    centered_text(ctx, WIDTH / 2, HEIGHT / 2, str(card['level']))
    ctx.set_source_rgb(0, 0, 0) # black
    ctx.set_font_size(24)
    multi_line_text(ctx, WIDTH / 2, HEIGHT / 2, card['term'], WIDTH - 3 * INSET)

# paints the back of the card
def make_back(ctx, card):
    background(ctx, (0.9, 1, 0.9)) # light green
    ctx.set_source_rgb(0.7, 1, 0.7) # somewhat darker green
    ctx.select_font_face('sans', cairo.FONT_SLANT_NORMAL, cairo.FONT_WEIGHT_NORMAL)
    ctx.set_font_size(144)
    centered_text(ctx, WIDTH / 2, HEIGHT / 2, str(card['level']))
    ctx.set_source_rgb(0, 0, 0) # black
    ctx.set_font_size(12)
    multi_line_text(ctx, WIDTH / 2, HEIGHT / 2, card['definition'], WIDTH - 3 * INSET)

# displays a number in the bottom left corner of the card
def number(ctx, i):
    ctx.set_source_rgb(0, 0, 0) # black
    ctx.set_font_size(6)
    ctx.move_to(1.5 * INSET, HEIGHT - 1.5 * INSET)
    ctx.show_text(str(i))

# paints a black dot at coordinates (x, y)
def black_dot(ctx, x, y):
    ctx.set_source_rgb(0, 0, 0) # black
    ctx.new_path()
    ctx.arc(x, y, DOT_RADIUS, 0, 2 * pi)
    ctx.fill()

# draw all cards. cutting_lines may be YES (draw card and cutting lines), NO (draw only card)
# or ONLY (draw only cuttong lines)
def draw_cards(cards, surface, cutting_lines):
    front = True # first draw a page of card fronts
    i = 0
    while i < len(cards):
        if front:
            # determine position, left to right, top to bottom
            card_in_row = i % CARDS_HORI
            x = MARGIN + card_in_row * WIDTH
            if card_in_row == 1: # make room for black dots
                x += 2 * DOT_RADIUS
            elif card_in_row > 1:
                x += 4 * DOT_RADIUS
            y = MARGIN + ((i // CARDS_HORI) % CARDS_VERTI) * HEIGHT
            sub_surface = surface.create_for_rectangle(x, y, WIDTH, HEIGHT)
            ctx = cairo.Context(sub_surface)
            if cutting_lines:
                cutting_line(ctx)
            if cutting_lines != ONLY:
                make_front(ctx, cards[i])
                number(ctx, i + 1)
            if (i + 1) % CARDS_PER_PAGE == 0 or i == len(cards) - 1: # page full or final page?
                if DOTS:
                    ctx = cairo.Context(surface)
                    black_dot(ctx, MARGIN + WIDTH + DOT_RADIUS, MARGIN + HEIGHT)
                    black_dot(ctx, MARGIN + 2 * WIDTH + 3 * DOT_RADIUS, MARGIN + 2 * HEIGHT)
                    black_dot(ctx, MARGIN + WIDTH + DOT_RADIUS, MARGIN + 3 * HEIGHT)
                surface.show_page()
            if cutting_lines != ONLY:
                if (i + 1) % CARDS_PER_PAGE == 0: # page full?
                    front = False # switch to drawing card backs
                    i -= CARDS_PER_PAGE - 1 # reset counter to start of page
                elif i == len(cards) - 1: # last card when page is not yet full?
                    front = False # switch to drawing card backs
                    i -= i % CARDS_PER_PAGE # reset counter to start of page
        if not front: # do not use 'else' here; both cases are executed at the turn of a page
            # determine position, right to left, top to bottom
            card_in_row = i % CARDS_HORI
            x = round(PAGE_WIDTH - MARGIN - (card_in_row + 1) * WIDTH) # use round because PAGE_WITH is likely not an integer
            if card_in_row == 1: # make room for black dots
                x -= 2 * DOT_RADIUS
            elif card_in_row > 1:
                x -= 4 * DOT_RADIUS
            y = MARGIN + ((i // CARDS_HORI) % CARDS_VERTI) * HEIGHT
            sub_surface = surface.create_for_rectangle(x, y, WIDTH, HEIGHT)
            ctx = cairo.Context(sub_surface)
            if cutting_lines:
                cutting_line(ctx)
            make_back(ctx, cards[i])
            number(ctx, i + 1)
            if (i + 1) % CARDS_PER_PAGE == 0: # page full?
                surface.show_page()
                front = True # switch back to drawing card fronts
        i += 1

# load entire input file
cards = []
wb = openpyxl.load_workbook(INPUT_FILE)
ws = wb.active
for row in ws.iter_rows(min_row=2, values_only=True): # skip header row
    cards.append({'level': row[0], 'term': row[1], 'definition': row[2]})

# create output file
surface = cairo.PDFSurface(OUTPUT_FILE, PAGE_WIDTH, PAGE_HEIGHT)
draw_cards(cards, surface, NO if SEPARATE_CUTTING_LINES else YES)
surface.finish()

if SEPARATE_CUTTING_LINES:
    # create cutting lines
    surface = cairo.PDFSurface(CUTTING_FILE, PAGE_WIDTH, PAGE_HEIGHT)
    draw_cards(cards, surface, ONLY)
    surface.finish()
