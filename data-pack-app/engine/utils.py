"""
Utility functions for Excel formula generation.
Column math, formula builders, and layout computation.
"""
from openpyxl.utils import get_column_letter, column_index_from_string


def col_letter(n):
    """Convert 1-based column number to Excel column letter. e.g. 1='A', 27='AA'."""
    return get_column_letter(n)


def col_num(letter):
    """Convert Excel column letter to 1-based number. e.g. 'A'=1, 'AA'=27."""
    return column_index_from_string(letter)


def make_range(sheet, col_letter_str, first_row, last_row, abs_col=True, abs_row=True):
    """Build a range reference like 'Sheet'!$A$2:$A$23498."""
    dc = '$' if abs_col else ''
    dr = '$' if abs_row else ''
    return f"'{sheet}'!{dc}{col_letter_str}{dr}{first_row}:{dc}{col_letter_str}{dr}{last_row}"


def make_cell(sheet, col_letter_str, row, abs_col=True, abs_row=True):
    """Build a cell reference like 'Sheet'!$A$1."""
    dc = '$' if abs_col else ''
    dr = '$' if abs_row else ''
    return f"'{sheet}'!{dc}{col_letter_str}{dr}{row}"


def sumifs(sum_range, *criteria_pairs):
    """Build a SUMIFS formula. criteria_pairs: [(range, criteria), ...]."""
    parts = [sum_range]
    for cr_range, cr_value in criteria_pairs:
        parts.extend([cr_range, cr_value])
    return f"=SUMIFS({','.join(parts)})"


def countifs(*criteria_pairs):
    """Build a COUNTIFS formula. criteria_pairs: [(range, criteria), ...]."""
    parts = []
    for cr_range, cr_value in criteria_pairs:
        parts.extend([cr_range, cr_value])
    return f"=COUNTIFS({','.join(parts)})"


def compute_clean_layout(num_attrs, num_dates, yoy_offset):
    """
    Compute column positions for a clean data tab.

    Returns a dict with all column positions (1-based column numbers).

    Layout:
    B          = Customer ID
    C..C+n-1   = Attribute columns
    next       = Cohort
    next       = Rank
    next       = Label column (for "Quarter"/"Year" in rows 2-3)
    next       = First ARR date column
    ...        = ARR date columns
    gap        = separator
    next..     = Churn section
    gap        = separator
    next..     = Downsell section
    gap        = separator
    next..     = Upsell section
    gap        = separator
    next..     = New Business section
    """
    num_derived = num_dates - yoy_offset

    cust_id_col = 2  # B
    attr_start = 3   # C
    attr_end = attr_start + num_attrs - 1  # last attribute column
    cohort_col = attr_end + 1
    rank_col = cohort_col + 1
    label_col = rank_col + 1  # H equivalent

    arr_start = label_col + 1  # I equivalent
    arr_end = arr_start + num_dates - 1

    churn_start = arr_end + 2  # +1 gap, +1 start
    churn_end = churn_start + num_derived - 1

    downsell_start = churn_end + 2
    downsell_end = downsell_start + num_derived - 1

    upsell_start = downsell_end + 2
    upsell_end = upsell_start + num_derived - 1

    new_biz_start = upsell_end + 2
    new_biz_end = new_biz_start + num_derived - 1

    return {
        'cust_id': cust_id_col,
        'attr_start': attr_start,
        'attr_end': attr_end,
        'num_attrs': num_attrs,
        'cohort': cohort_col,
        'rank': rank_col,
        'label': label_col,
        'arr_start': arr_start,
        'arr_end': arr_end,
        'churn_start': churn_start,
        'churn_end': churn_end,
        'downsell_start': downsell_start,
        'downsell_end': downsell_end,
        'upsell_start': upsell_start,
        'upsell_end': upsell_end,
        'new_biz_start': new_biz_start,
        'new_biz_end': new_biz_end,
        'yoy_offset': yoy_offset,
        'num_dates': num_dates,
        'num_derived': num_derived,
    }


def granularity_label(granularity):
    """Return display label for a granularity level."""
    return {
        'monthly': 'Monthly',
        'quarterly': 'Quarterly',
        'annual': 'Annual',
    }[granularity]


def cohort_label_formula(granularity, label_col_letter, quarter_row_val=None, year_row_val=None):
    """
    Return the formula pattern for cohort period labels.
    quarterly: ="Q"&B10&"'"&RIGHT(C10,2)
    annual: ="FY"&"'"&RIGHT(I3,2)
    """
    if granularity == 'quarterly':
        return '="Q"&{q}&"\'"&RIGHT({y},2)'
    elif granularity == 'annual':
        return '="FY"&"\'"&RIGHT({y},2)'
    else:  # monthly - dates are values, no formula needed for labels
        return None


def period_label_formula(granularity, quarter_cell, year_cell):
    """Build the period label formula for column headers (row 6)."""
    if granularity == 'quarterly':
        return f'="Q"&{quarter_cell}&"\'"&RIGHT({year_cell},2)'
    elif granularity == 'annual':
        return f'="FY"&"\'"&RIGHT({year_cell},2)'
    return None  # monthly uses date values


def get_yoy_offset(granularity):
    """Get the year-over-year offset for a given granularity."""
    return {'monthly': 12, 'quarterly': 4, 'annual': 1}[granularity]
