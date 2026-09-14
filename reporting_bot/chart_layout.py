"""Shared, data-neutral layout helpers for report images."""
from __future__ import annotations

import io
import math
from PIL import Image, ImageDraw


def fitted_text(draw, xy, text, font_factory, *, width, size=20, minimum=12,
                bold=False, fill="#173A5E"):
    """Fit using actual glyph widths (including Cyrillic), never character counts."""
    text = str(text)
    font = font_factory(size, bold)
    while size > minimum and draw.textlength(text, font=font) > width:
        size -= 1
        font = font_factory(size, bold)
    if draw.textlength(text, font=font) > width:
        while text and draw.textlength(text + "...", font=font) > width:
            text = text[:-1]
        text += "..."
    draw.text(xy, text, font=font, fill=fill)
    return text


def comparison_chart(title, subtitle, labels, current, comparison, *, font,
                     formatter, comparison_label, cards=(), show_comparison=True):
    """Fixed value column prevents maximum bars from covering numeric labels."""
    row_height = 104
    start_y = 412 if cards else 242
    height = max(720, start_y + max(1, len(labels)) * row_height + 86)
    image = Image.new("RGB", (1200, height), "#F3F6FA")
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((32, 24, 1168, height - 24), radius=28, fill="white")
    fitted_text(draw, (70, 52), title, font, width=1060, size=38, minimum=26, bold=True)
    fitted_text(draw, (70, 107), subtitle, font, width=1060, size=20, minimum=15, fill="#586A7E")
    if cards:
        for index, (label, value, detail) in enumerate(cards[:4]):
            left = 70 + index * 270
            draw.rounded_rectangle((left, 158, left + 250, 286), radius=15, fill="#EEF4FA")
            fitted_text(draw, (left + 16, 173), label, font, width=218, size=18)
            fitted_text(draw, (left + 16, 205), value, font, width=218, size=30, bold=True)
            fitted_text(draw, (left + 16, 250), detail, font, width=218, size=15, fill="#586A7E")
    legend_y = start_y - 70
    draw.rectangle((70, legend_y + 5, 89, legend_y + 21), fill="#255985")
    draw.text((100, legend_y), "Выбранный период", font=font(19), fill="#173A5E")
    if show_comparison:
        draw.rectangle((400, legend_y + 5, 419, legend_y + 21), fill="#E89A3C")
        fitted_text(draw, (430, legend_y), comparison_label, font, width=670, size=19)
    maximum = max([*current, *(comparison if show_comparison else ()), 1.0])
    bar_left, bar_right = 392, 924
    for index, label in enumerate(labels):
        y = start_y + index * row_height
        if index % 2 == 0:
            draw.rounded_rectangle((58, y - 10, 1140, y + 83), radius=12, fill="#F7F9FC")
        fitted_text(draw, (70, y + 15), label, font, width=302, size=20, minimum=14, bold=True)
        values = [(current[index], "#255985", y)]
        if show_comparison:
            values.append((comparison[index], "#E89A3C", y + 37))
        for value, color, top in values:
            draw.rounded_rectangle((bar_left, top, bar_right, top + 24), radius=6, fill="#E6EDF5")
            width = (bar_right - bar_left) * max(value, 0) / maximum
            if width > 0:
                draw.rounded_rectangle((bar_left, top, bar_left + width, top + 24), radius=min(6, width / 2), fill=color)
            fitted_text(draw, (944, top - 2), formatter(value), font, width=180, size=20, bold=True)
    if not labels:
        draw.text((70, start_y + 10), "Нет данных за выбранный период", font=font(26), fill="#586A7E")
    draw.text((70, height - 69), "EcoStar Reports · точные значения и детализация — в Excel", font=font(17), fill="#586A7E")
    output = io.BytesIO()
    image.save(output, format="PNG", optimize=True)
    return output.getvalue()


def daily_panel(draw, box, title, dates, launch, release, *, font, formatter):
    """Draw every day; use lines for dense periods, never silently aggregate."""
    left, top, right, bottom = box
    blue, orange = "#255985", "#D98726"
    draw.rounded_rectangle(box, radius=18, fill="#F8FAFD", outline="#DCE5EF", width=2)
    fitted_text(draw, (left + 22, top + 14), title, font, width=right-left-44, size=23, bold=True)
    for index, (label, color) in enumerate((("Запуск", blue), ("Выпуск", orange))):
        x = left + 24 + index * 150
        draw.rectangle((x, top + 51, x + 18, top + 65), fill=color)
        draw.text((x + 28, top + 45), label, font=font(17), fill="#586A7E")
    dense = len(dates) > 45
    if dense:
        fitted_text(draw, (left + 340, top + 46), "Каждая точка — день; без укрупнения", font,
                    width=right-left-365, size=15, fill="#586A7E")
    x0, x1 = left + 104, right - 32
    y0, y1 = top + 97, bottom - 44
    maximum = max([*launch, *release, 1.0])
    minimum = min([*launch, *release, 0.0])
    step = max(1, math.ceil((maximum - minimum) / 4))
    low = math.floor(minimum / step) * step
    high = math.ceil(maximum / step) * step
    def y(value):
        return y1 - (y1-y0) * (value-low) / (high-low)
    for value in range(low, high + 1, step):
        yy = y(value)
        draw.line((x0, yy, x1, yy), fill="#DFE7F0", width=1)
        label = formatter(value)
        fitted_text(draw, (left + 12, yy - 9), label, font, width=84, size=14, minimum=11, fill="#586A7E")
    if not dates or not any((*launch, *release)):
        draw.text((x0 + 15, y0 + 20), "Нет запусков и выпусков за период", font=font(21), fill="#586A7E")
    if not dates:
        return
    group_width = (x1-x0) / len(dates)
    if dense:
        for values, color in ((launch, blue), (release, orange)):
            points = [(x0 + group_width * (i + .5), y(value)) for i, value in enumerate(values)]
            draw.line(points, fill=color, width=5 if color == blue else 2)
    else:
        bar_width = min(22, group_width * .32)
        for index in range(len(dates)):
            center = x0 + group_width * (index + .5)
            for value, x, color in ((launch[index], center-bar_width-1, blue), (release[index], center+1, orange)):
                if not value:
                    continue
                draw.rectangle((x, min(y(0), y(value)), x+bar_width, max(y(0), y(value))), fill=color)
                label = formatter(value)
                # Labels must fit inside their own half of the day, including both series.
                if draw.textlength(label, font=font(14, True)) + 8 < group_width / 2:
                    draw.text((x, y(value)-20 if value > 0 else y(value)+4), label, font=font(14, True), fill=color)
    indices = list(range(0, len(dates), max(1, math.ceil(len(dates) / max(1, int((x1-x0)/85))))))
    if indices[-1] != len(dates)-1:
        if (len(dates)-1-indices[-1])*group_width < 75:
            indices.pop()
        indices.append(len(dates)-1)
    for index in indices:
        crosses_year = dates[0].year != dates[-1].year
        label = dates[index].strftime("%d.%m.%y" if crosses_year else "%d.%m")
        label_font = font(12 if crosses_year else 14)
        width = draw.textlength(label, font=label_font)
        draw.text((x0+group_width*(index+.5)-width/2, y1+12), label, font=label_font, fill="#586A7E")
