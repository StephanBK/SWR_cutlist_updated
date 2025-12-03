import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

# Conversion constants
inches_to_mm = 25.4
mm_to_inches = 1 / inches_to_mm
sq_inches_to_sq_feet = 1 / 144

# Display logo and title
st.image("ilogo.png", width=200)
st.title("SWR Cutlist")

# Project details inputs
project_name = st.text_input("Enter Project Name")
project_number = st.text_input("Enter Project Number", value="INO-")
prepared_by = st.text_input("Prepared By")

# System type and finish
system_type = st.selectbox("Select System Type", ["SWR-IG", "SWR-VIG", "SWR", "Custom"])
finish = st.selectbox("Select Finish", ["Mil Finish", "Clear Anodized", "Black Anodized", "Painted"])

# Determine Glass Offset and profile number
offset_unit = 'mm'
if system_type == 'SWR-IG':
    glass_offset = 11.1125
    profile_number = '03003'
elif system_type == 'SWR-VIG':
    glass_offset = 11.1125
    profile_number = '03004'
elif system_type == 'SWR':
    # Let the user choose between the two SWR profiles
    swr_profile_choice = st.radio(
        "Select SWR Profile (Profile – Glass Offset)",
        [
            "03002 – 7.571 mm glass offset",
            "03111 – 11.11 mm glass offset",
        ]
    )
    if swr_profile_choice.startswith("03002"):
        profile_number = '03002'
        glass_offset = 7.571
    else:
        profile_number = '03111'
        glass_offset = 11.11
else:
    unit_offset = st.radio("Select Unit for Glass Offset", ["Inches", "Millimeters"], index=0)
    if unit_offset == 'Inches':
        glass_offset = st.number_input("Enter Glass Offset (in inches)", value=0.0) * inches_to_mm
        offset_unit = 'inches'
    else:
        glass_offset = st.number_input("Enter Glass Offset (in mm)", value=0.0)
        offset_unit = 'mm'

if system_type != 'Custom':
    st.write(f"Using a Glass Offset of {glass_offset:.3f} mm for system {system_type}")

# Input parameters and units
unit_tolerance = st.radio("Unit for Glass Cutting Tolerance", ["Inches", "Millimeters"], index=0)
tol_unit = 'inches' if unit_tolerance == 'Inches' else 'mm'
if unit_tolerance == 'Inches':
    glass_cutting_tolerance = st.number_input("Enter Glass Cutting Tolerance (in inches)", value=0.0625, format="%.4f")
else:
    val_mm = st.number_input("Enter Glass Cutting Tolerance (in mm)", value=0.0625 * inches_to_mm, format="%.3f")
    glass_cutting_tolerance = val_mm * mm_to_inches

unit_joint_top = st.radio("Unit for Joint Top", ["Inches", "Millimeters"], index=0)
top_unit = 'inches' if unit_joint_top == 'Inches' else 'mm'
if unit_joint_top == 'Inches':
    joint_top = st.number_input("Enter Joint Top (in inches)", value=0.5, format="%.3f")
else:
    val_mm = st.number_input("Enter Joint Top (in mm)", value=0.5 * inches_to_mm, format="%.3f")
    joint_top = val_mm * mm_to_inches

unit_joint_bottom = st.radio("Unit for Joint Bottom", ["Inches", "Millimeters"], index=0)
bottom_unit = 'inches' if unit_joint_bottom == 'Inches' else 'mm'
if unit_joint_bottom == 'Inches':
    joint_bottom = st.number_input("Enter Joint Bottom (in inches)", value=0.125, format="%.3f")
else:
    val_mm = st.number_input("Enter Joint Bottom (in mm)", value=0.125 * inches_to_mm, format="%.3f")
    joint_bottom = val_mm * mm_to_inches

unit_joint_left = st.radio("Unit for Joint Left", ["Inches", "Millimeters"], index=0)
left_unit = 'inches' if unit_joint_left == 'Inches' else 'mm'
if unit_joint_left == 'Inches':
    joint_left = st.number_input("Enter Joint Left (in inches)", value=0.25, format="%.3f")
else:
    val_mm = st.number_input("Enter Joint Left (in mm)", value=0.25 * inches_to_mm, format="%.3f")
    joint_left = val_mm * mm_to_inches

unit_joint_right = st.radio("Unit for Joint Right", ["Inches", "Millimeters"], index=0)
right_unit = 'inches' if unit_joint_right == 'Inches' else 'mm'
if unit_joint_right == 'Inches':
    joint_right = st.number_input("Enter Joint Right (in inches)", value=0.25, format="%.3f")
else:
    val_mm = st.number_input("Enter Joint Right (in mm)", value=0.25 * inches_to_mm, format="%.3f")
    joint_right = val_mm * mm_to_inches

# Determine part number
part_number = f"{system_type}-{profile_number}" if system_type != 'Custom' else 'Custom'

# Download template
template_path = 'SWR template.csv'
st.download_button('Download Template', data=open(template_path, 'rb').read(), file_name='SWR_template.csv', mime='text/csv')

# File upload
uploaded_file = st.file_uploader('Upload a CSV file', type='csv')

# Prepare parameters rows
params = [
    ('Prepared By', prepared_by, ''),
    ('System Type', system_type, ''),
    ('Finish', finish, ''),
    ('Glass Offset', glass_offset, offset_unit),
    ('Glass Cutting Tolerance', glass_cutting_tolerance, tol_unit),
    ('Joint Top', joint_top, top_unit),
    ('Joint Bottom', joint_bottom, bottom_unit),
    ('Joint Left', joint_left, left_unit),
    ('Joint Right', joint_right, right_unit)
]

# Process when file uploaded
if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.dataframe(df)

    # Convert dims
    df['Overall Width mm'] = df['Overall Width in'] * inches_to_mm
    df['Overall Height mm'] = df['Overall Height in'] * inches_to_mm
    j_l = joint_left * inches_to_mm
    j_r = joint_right * inches_to_mm
    j_t = joint_top * inches_to_mm
    j_b = joint_bottom * inches_to_mm

    # SWR dims
    df['SWR Width mm'] = df['Overall Width mm'] - j_l - j_r
    df['SWR Height mm'] = df['Overall Height mm'] - j_t - j_b
    df['SWR Width in'] = df['SWR Width mm'] * mm_to_inches
    df['SWR Height in'] = df['SWR Height mm'] * mm_to_inches

    # Glass dims
    df['Glass Width mm'] = df['SWR Width mm'] - 2 * glass_offset
    df['Glass Height mm'] = df['SWR Height mm'] - 2 * glass_offset
    df['Glass Width in'] = df['Glass Width mm'] * mm_to_inches
    df['Glass Height in'] = df['Glass Height mm'] * mm_to_inches

    # Helper: sixteenth rounding
    def to_sixteenth(x):
        total = round(x * 16)
        w, f = divmod(total, 16)
        return f"{w} {f}/16" if f else f"{w}"

    # --- Glass File Export ---
    glass_df = pd.DataFrame({
        'Item': range(1, len(df) + 1),
        'Glass Width in': df['Glass Width in'],
        'Glass Width (1/16)': df['Glass Width in'].apply(to_sixteenth),
        'Glass Height in': df['Glass Height in'],
        'Glass Height (1/16)': df['Glass Height in'].apply(to_sixteenth),
        'Area Each (ft²)': (df['Glass Width in'] * df['Glass Height in']) * sq_inches_to_sq_feet,
        'Qty': df['Qty'],
        'Area Total (ft²)': df['Qty'] * (df['Glass Width in'] * df['Glass Height in']) * sq_inches_to_sq_feet
    })
    # Add totals row
    totals = pd.DataFrame([{col: (glass_df[col].sum() if col in ['Qty', 'Area Total (ft²)'] else None) for col in glass_df.columns}])
    totals.at[0, 'Item'] = 'Totals'
    glass_df = pd.concat([glass_df, totals], ignore_index=True)

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        ws = writer.book.add_worksheet('Glass')
        ws.insert_image('A1', 'ilogo.png', {'x_scale': 0.2, 'y_scale': 0.2})
        ws.write_row('A7', ['Project Name:', project_name])
        ws.write_row('A8', ['Project Number:', project_number])
        ws.write_row('A9', ['Date Created:', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
        ws.write_row('A10', ['Prepared By:', prepared_by])
        glass_df.to_excel(writer, sheet_name='Glass', startrow=12, index=False)
        pws = writer.book.add_worksheet('Parameters')
        for idx, (lbl, val, unit) in enumerate(params, start=1):
            pws.write(idx - 1, 0, lbl)
            pws.write(idx - 1, 1, val)
            pws.write(idx - 1, 2, unit)
    st.download_button('Download Glass File', buf.getvalue(), file_name=f'INO_{project_number}_SWR_Glass.xlsx')

    # --- AggCutOnly File Export ---
    df['Qty x 2'] = df['Qty'] * 2
    wc = df.groupby('SWR Width in')['Qty'].sum().sort_values(ascending=False)
    hc = df.groupby('SWR Height in')['Qty'].sum().sort_values(ascending=False)
    dims = pd.Index(list(wc.index) + list(hc.index)).unique()
    agg = pd.DataFrame(0, index=dims, columns=['Part #', 'Miter'] + list(df['Tag'].unique()) + ['Total QTY'])
    agg['Part #'] = part_number
    agg['Miter'] = '**'
    for _, r in df.iterrows():
        w, h, t, q2 = r['SWR Width in'], r['SWR Height in'], r['Tag'], r['Qty x 2']
        agg.at[w, t] += q2
        agg.at[h, t] += q2
    agg['Total QTY'] = agg[list(df['Tag'].unique())].sum(axis=1)
    agg = agg.reset_index().rename(columns={'index': 'Finished Length in'})

    buf2 = BytesIO()
    with pd.ExcelWriter(buf2, engine='xlsxwriter') as writer:
        ws = writer.book.add_worksheet('AggCutOnly')
        ws.insert_image('A1', 'ilogo.png', {'x_scale': 0.2, 'y_scale': 0.2})
        ws.write_row('A7', ['Project Name:', project_name])
        ws.write_row('A8', ['Project Number:', project_number])
        ws.write_row('A9', ['Date Created:', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
        ws.write_row('A10', ['Prepared By:', prepared_by])
        ws.write_row('A11', ['Finish:', finish])
        agg.to_excel(writer, sheet_name='AggCutOnly', startrow=12, index=False)
        pws = writer.book.add_worksheet('Parameters')
        for idx, (lbl, val, unit) in enumerate(params, start=1):
            pws.write(idx - 1, 0, lbl)
            pws.write(idx - 1, 1, val)
            pws.write(idx - 1, 2, unit)
    st.download_button('Download AggCutOnly File', buf2.getvalue(), file_name=f'INO_{project_number}_SWR_AggCutOnly.xlsx')

    # --- TagDetails File Export ---
    buf3 = BytesIO()
    with pd.ExcelWriter(buf3, engine='xlsxwriter') as writer:
        for tag in df['Tag'].unique():
            rows = {'Item': [], 'Position': [], 'Quantity': [], 'Length (mm)': [], 'Length (in)': []}
            subset = df[df['Tag'] == tag]
            for idx, r in subset.iterrows():
                for pos, length in [('left', r['SWR Width mm']), ('right', r['SWR Width mm']),
                                    ('top', r['SWR Height mm']), ('bottom', r['SWR Height mm'])]:
                    rows['Item'].append(idx + 1)
                    rows['Position'].append(pos)
                    rows['Quantity'].append(r['Qty'] * 2)
                    rows['Length (mm)'].append(length)
                    rows['Length (in)'].append(length * mm_to_inches)
            tag_df = pd.DataFrame(rows)
            ws = writer.book.add_worksheet(str(tag))
            ws.insert_image('A1', 'ilogo.png', {'x_scale': 0.2, 'y_scale': 0.2})
            ws.write_row('A7', ['Project Name:', project_name])
            ws.write_row('A8', ['Project Number:', project_number])
            ws.write_row('A9', ['Date Created:', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
            ws.write_row('A10', ['Prepared By:', prepared_by])
            tag_df.to_excel(writer, sheet_name=str(tag), startrow=12, index=False)
        pws = writer.book.add_worksheet('Parameters')
        for idx, (lbl, val, unit) in enumerate(params, start=1):
            pws.write(idx - 1, 0, lbl)
            pws.write(idx - 1, 1, val)
            pws.write(idx - 1, 2, unit)
    st.download_button('Download TagDetails File', buf3.getvalue(), file_name=f'INO_{project_number}_SWR_TagDetails.xlsx')

    # --- SWR Table File Export ---
    buf4 = BytesIO()
    with pd.ExcelWriter(buf4, engine='xlsxwriter') as writer:
        ws = writer.book.add_worksheet('Table')
        ws.insert_image('A1', 'ilogo.png', {'x_scale': 0.2, 'y_scale': 0.2})
        ws.write_row('A7', ['Project Name:', project_name])
        ws.write_row('A8', ['Project Number:', project_number])
        ws.write_row('A9', ['Date Created:', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
        ws.write_row('A10', ['Prepared By:', prepared_by])
        ws.write_row('A11', ['Finish:', finish])
        ws.write_row('A12', ['Glass Cutting Tolerance:', glass_cutting_tolerance])
        ws.write_row('A13', ['Tolerance Unit:', tol_unit])
        ws.write_row('A14', ['Joint Top:', joint_top])
        ws.write_row('A15', ['Joint Top Unit:', top_unit])
        ws.write_row('A16', ['Joint Bottom:', joint_bottom])
        ws.write_row('A17', ['Joint Bottom Unit:', bottom_unit])
        ws.write_row('A18', ['Joint Left:', joint_left])
        ws.write_row('A19', ['Joint Left Unit:', left_unit])
        ws.write_row('A20', ['Joint Right:', joint_right])
        ws.write_row('A21', ['Joint Right Unit:', right_unit])
        df.drop(columns=['Qty x 2'], errors='ignore').to_excel(writer, sheet_name='Table', startrow=23, index=False)
        pws = writer.book.add_worksheet('Parameters')
        for idx, (lbl, val, unit) in enumerate(params, start=1):
            pws.write(idx - 1, 0, lbl)
            pws.write(idx - 1, 1, val)
            pws.write(idx - 1, 2, unit)
    st.download_button('Download SWR Table File', buf4.getvalue(), file_name=f'INO_{project_number}_SWR_Table.xlsx')