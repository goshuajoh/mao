"""
Signify Order Processor - Web App v2.0
=======================================
Streamlit web interface for non-technical users
Deploy to Streamlit Cloud for free!

Changes in v2.0:
- POAF column formatted as whole number (no scientific notation)
- OrderHub Tax Rate: 13 (not 0.13)
- OrderHub Solution: 'N/A'
- Product suffix only for specific products (ESP32-SOLO-1, ESP32-C3-MINI-1, ESP-WROOM-02D)
- New export file with 格式转换 filled from order history
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io

def check_password():
    """Returns True if user enters correct password"""
    def password_entered():
        if st.session_state["password"] == "Ma0!":
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store password
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        # Password incorrect
        st.text_input(
            "Password", type="password", on_change=password_entered, key="password"
        )
        st.error("😕 Password incorrect")
        return False
    else:
        # Password correct
        return True

if not check_password():
    st.stop()


# Page config
st.set_page_config(
    page_title="Signify Order Processor",
    page_icon="📦",
    layout="wide"
)

# Title
st.title("📦 Signify Order Processor")
st.markdown("Upload your export file and get processed orders in seconds!")

# Important note about downloads
with st.expander("💡 Quick Tips", expanded=False):
    st.markdown("""
    **After downloading files:**
    - Files are saved to your computer automatically
    - Page may reload - this is normal!
    - Your downloads are ready in your Downloads folder
    
    **To update history:**
    1. Download your output files first
    2. Then click "Generate Updated History"
    3. Download the updated history file
    4. Use it next time to avoid duplicates!
    """)

# Sidebar for reference files
with st.sidebar:
    st.header("📁 Reference Files")
    
    # Tab for normal user vs admin
    tab_user, tab_admin = st.tabs(["📤 Upload", "⚙️ Admin"])
    
    with tab_user:
        history_file = st.file_uploader(
            "Order History (signify_order_list.xlsx)",
            type=['xlsx'],
            key='history',
            help="Upload your order history file"
        )
        
        master_file = st.file_uploader(
            "Master Reference (Signify_Master_Reference_COMPLETE.xlsx)",
            type=['xlsx'],
            key='master',
            help="Upload your master reference file"
        )
        
        st.markdown("---")
        st.markdown("### ℹ️ Instructions")
        st.markdown("""
        1. Upload reference files (sidebar) - **once only**
        2. Upload today's export file (main area)
        3. Click 'Process Orders'
        4. Download your output files!
        """)
    
    with tab_admin:
        st.markdown("### 🔧 Update Reference Files")
        st.markdown("For developers/admins only")
        
        admin_password = st.text_input("Admin Password:", type="password", key="admin_pw")
        
        if admin_password == "signify_admin_2024":  # Change this password!
            st.success("✅ Admin access granted")
            
            st.markdown("---")
            st.markdown("#### 📝 Update Order History")
            st.markdown("Add newly processed orders to history")
            
            if 'df_new' in st.session_state and st.session_state.get('processing_complete'):
                st.info(f"✓ {len(st.session_state.df_new)} orders ready to add")
                
                if st.button("➕ Add to Order History", use_container_width=True):
                    # This will be implemented after processing
                    st.session_state.update_history = True
                    st.success("✓ Orders will be added after download")
            else:
                st.warning("⚠️ Process orders first, then update history")
            
            st.markdown("---")
            st.markdown("#### 🔄 Upload New Reference Files")
            st.markdown("Replace existing reference files")
            
            new_history = st.file_uploader(
                "New Order History File",
                type=['xlsx'],
                key='new_history',
                help="Completely replace order history"
            )
            
            new_master = st.file_uploader(
                "New Master Reference File",
                type=['xlsx'],
                key='new_master',
                help="Completely replace master reference"
            )
            
            if new_history:
                st.download_button(
                    "⬇️ Download Updated History",
                    data=new_history.getvalue(),
                    file_name="signify_order_list_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                st.info("💡 Save this file and use it for next session")
            
            if new_master:
                st.download_button(
                    "⬇️ Download Updated Master",
                    data=new_master.getvalue(),
                    file_name="Signify_Master_Reference_COMPLETE_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                st.info("💡 Save this file and use it for next session")
        
        elif admin_password:
            st.error("❌ Incorrect password")

# Main area - Export file upload
st.header("1️⃣ Upload Export File")
export_file = st.file_uploader(
    "Upload today's export file from customer",
    type=['xlsx'],
    key='export',
    help="The daily export file you receive from customer each day"
)

# Check if all files are uploaded
if export_file and history_file and master_file:
    st.success("✅ All files uploaded!")
    
    # Process button
    if st.button("🚀 Process Orders", type="primary", use_container_width=True):
        
        # Progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            # Step 1: Load files
            status_text.text("Step 1/12: Loading files...")
            progress_bar.progress(8)
            
            df_export = pd.read_excel(export_file)
            df_history = pd.read_excel(history_file, sheet_name='1.6-now')
            
            st.write(f"📊 Loaded {len(df_export)} orders from export")
            
            # Step 2: Parse Product Desc
            status_text.text("Step 2/12: Parsing product descriptions...")
            progress_bar.progress(16)
            
            split_cols = df_export['Product Desc.'].str.split('_', expand=True)
            df_export['Product_Base'] = split_cols[0]
            df_export['Brand'] = split_cols[1]
            df_export['PN'] = split_cols[2]
            
            # Step 3: Create UUID
            status_text.text("Step 3/12: Creating unique IDs...")
            progress_bar.progress(24)
            
            df_export['UUID'] = (
                df_export['PO No.'].astype(str).str.strip() + 
                df_export['PN'].astype(str).str.strip() +
                df_export['Qty'].astype(float).astype(int).astype(str)
            )
            
            # Step 4: Deduplication
            status_text.text("Step 4/12: Checking for new orders...")
            progress_bar.progress(32)
            
            # UUID is in column index 19 (合并)
            uuid_col = df_history.columns[19]
            existing_uuids = set(df_history[uuid_col].dropna().astype(str))
            
            df_export['Is_New'] = ~df_export['UUID'].isin(existing_uuids)
            new_count = df_export['Is_New'].sum()
            existing_count = len(df_export) - new_count
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Orders", len(df_export))
            col2.metric("NEW Orders", new_count, delta=f"+{new_count}")
            col3.metric("Already Processed", existing_count)
            
            if new_count == 0:
                st.warning("⚠️ No new orders to process! All orders have been processed before.")
                st.stop()
            
            df_new = df_export[df_export['Is_New']].copy()
            
            # Step 5: Prepare data
            status_text.text("Step 5/12: Preparing order data...")
            progress_bar.progress(40)
            
            df_new['Unit_Price'] = df_new['Price'] / df_new['PrU.']
            df_new['Ship_To_Clean'] = df_new['Ship-To Loc.'].astype(str).str.lstrip('0')
            df_new['Ship_To_Clean'] = pd.to_numeric(df_new['Ship_To_Clean'], errors='coerce').fillna(0).astype(int)
            
            df_new['Is_Matter'] = df_new['PN'].astype(str).str.endswith('M')
            df_new['PN_Clean'] = df_new.apply(
                lambda row: row['PN'][:-1] if row['Is_Matter'] and isinstance(row['PN'], str) else row['PN'],
                axis=1
            )
            
            matter_count = df_new['Is_Matter'].sum()
            non_matter_count = len(df_new) - matter_count
            
            col1, col2 = st.columns(2)
            col1.metric("Matter Products", matter_count)
            col2.metric("Non-Matter Products", non_matter_count)
            
            # Step 6: Load firmware tables from master file
            status_text.text("Step 6/12: Loading firmware data...")
            progress_bar.progress(48)
            
            df_non_matter = pd.read_excel(master_file, sheet_name='NonMatter_Firmware')
            df_matter = pd.read_excel(master_file, sheet_name='Matter_Firmware')
            
            # Step 7: Firmware lookup
            status_text.text("Step 7/12: Looking up firmware versions...")
            progress_bar.progress(56)
            
            def lookup_non_matter_firmware(pn_clean, product_base):
                """
                Search for PN as a value within any column (not as column header).
                When found, return the firmware from column 4 of that row.
                """
                pn_str = str(pn_clean)
                
                for col_idx in range(5, len(df_non_matter.columns)):
                    col = df_non_matter.columns[col_idx]
                    for idx, value in df_non_matter[col].items():
                        if pd.notna(value):
                            if str(value) == pn_str or (isinstance(value, (int, float)) and str(int(value)) == pn_str):
                                firmware = df_non_matter.iloc[idx, 4]
                                if pd.notna(firmware) and str(firmware).startswith('V'):
                                    return firmware
                return None
            
            def lookup_matter_firmware(pn_clean, product_base):
                """
                Search for PN as a value within any column (not as column header).
                When found, return the firmware from column 4 of that row.
                """
                try:
                    pn_numeric = int(float(pn_clean))
                except:
                    pn_numeric = None
                
                pn_str = str(pn_clean)
                
                for col_idx in range(5, len(df_matter.columns)):
                    col = df_matter.columns[col_idx]
                    for idx, value in df_matter[col].items():
                        if pd.notna(value):
                            match = False
                            if isinstance(value, (int, float)):
                                if pn_numeric is not None and int(value) == pn_numeric:
                                    match = True
                            elif str(value) == pn_str:
                                match = True
                            
                            if match:
                                firmware = df_matter.iloc[idx, 4]
                                if pd.notna(firmware) and str(firmware).startswith('V'):
                                    return firmware
                return None
            
            def lookup_firmware_from_history(pn_clean):
                """Fallback: Look up firmware from order history for this PN"""
                matches = df_history[
                    df_history[df_history.columns[5]].astype(str).str.contains(str(pn_clean), case=False, na=False)
                ]
                if len(matches) > 0:
                    fw_values = matches[matches.columns[6]].dropna()
                    valid_fw = fw_values[fw_values.astype(str).str.startswith('V')]
                    if len(valid_fw) > 0:
                        return valid_fw.mode()[0] if len(valid_fw.mode()) > 0 else valid_fw.iloc[-1]
                return None
            
            df_new['Firmware'] = None
            for idx, row in df_new.iterrows():
                fw = None
                if row['Is_Matter']:
                    fw = lookup_matter_firmware(row['PN_Clean'], row['Product_Base'])
                else:
                    fw = lookup_non_matter_firmware(row['PN_Clean'], row['Product_Base'])
                
                if fw is None:
                    fw = lookup_firmware_from_history(row['PN_Clean'])
                
                df_new.at[idx, 'Firmware'] = fw
            
            firmware_found = df_new['Firmware'].notna().sum()
            firmware_missing = df_new['Firmware'].isna().sum()
            
            if firmware_missing > 0:
                st.warning(f"⚠️ {firmware_missing} orders missing firmware (will be empty in output)")
            
            # Step 8: Add product suffixes (UPDATED LOGIC)
            status_text.text("Step 8/12: Normalizing product names...")
            progress_bar.progress(64)
            
            def add_product_suffix(product_base):
                """
                Only add suffixes to specific products:
                - ESP32-SOLO-1 → ESP32-SOLO-1-H4
                - ESP32-C3-MINI-1 → ESP32-C3-MINI-1-H4
                - ESP-WROOM-02D → ESP-WROOM-02D-H2
                
                Do NOT add suffix to: SIG149, WIZ149, WIZ2012, etc.
                """
                if pd.isna(product_base):
                    return product_base
                product_base = str(product_base)
                
                # Only add -H4 to ESP32-SOLO-1 (but not if already has -H4)
                if 'ESP32-SOLO-1' in product_base and '-H4' not in product_base:
                    # Make sure it's not a SIG or WIZ variant
                    if 'SIG' not in product_base and 'WIZ' not in product_base:
                        return product_base + '-H4'
                
                # Only add -H4 to ESP32-C3-MINI-1 (but not if already has -H4)
                if 'ESP32-C3-MINI-1' in product_base and '-H4' not in product_base:
                    # Make sure it's not a SIG or WIZ variant
                    if 'SIG' not in product_base and 'WIZ' not in product_base:
                        return product_base + '-H4'
                
                # Only add -H2 to ESP-WROOM-02D (but not if already has -H2)
                if 'ESP-WROOM-02D' in product_base and '-H2' not in product_base:
                    return product_base + '-H2'
                
                # For everything else (including SIG149, WIZ149, WIZ2012), return as-is
                return product_base
            
            df_new['Product_MPN'] = df_new['Product_Base'].apply(add_product_suffix)
            
            # Step 9: Load ship-to data
            status_text.text("Step 9/12: Loading ship-to locations...")
            progress_bar.progress(72)
            
            df_shipto = pd.read_excel(master_file, sheet_name='ShipTo_Locations')
            df_shipto['Ship_To_Code'] = df_shipto['Ship_To_Code'].astype(str)
            
            df_new['Ship_To_Code_Str'] = df_new['Ship_To_Clean'].astype(str)
            df_new = df_new.merge(
                df_shipto[['Ship_To_Code', 'Customer_Name', 'Full_Address']],
                left_on='Ship_To_Code_Str',
                right_on='Ship_To_Code',
                how='left'
            )
            
            # Step 9.5: Load Client Ref mapping
            status_text.text("Step 9/12: Looking up Client References...")
            progress_bar.progress(76)
            
            try:
                df_client_ref = pd.read_excel(master_file, sheet_name='Client_Ref_Mapping')
                client_ref_dict = {}
                
                for idx, row in df_client_ref.iterrows():
                    if pd.notna(row.get('Product_Code')) and pd.notna(row.get('Customer_Code')):
                        lookup_key = str(int(row['Product_Code'])) + str(int(row['Customer_Code']))
                        client_ref = row.get('Customer_Material_Number', '')
                        if pd.notna(client_ref):
                            client_ref_dict[lookup_key] = client_ref
                
                def lookup_client_ref(product, ship_to_clean):
                    lookup_key = str(int(product)) + str(int(ship_to_clean))
                    return client_ref_dict.get(lookup_key, '')
                
                df_new['Client_Ref'] = df_new.apply(
                    lambda row: lookup_client_ref(row['Product'], row['Ship_To_Clean']) 
                    if pd.notna(row['Product']) and pd.notna(row['Ship_To_Clean']) else '',
                    axis=1
                )
                
                client_ref_found = (df_new['Client_Ref'] != '').sum()
                if client_ref_found > 0:
                    st.info(f"ℹ️ Found {client_ref_found} Client References from mapping table")
            except Exception as e:
                st.warning(f"⚠️ Could not load Client Ref mapping: {e}")
                df_new['Client_Ref'] = ''
            
            # Step 10: Generate factory output
            status_text.text("Step 10/12: Generating factory template...")
            progress_bar.progress(82)
            
            # Convert POAF to integer to avoid scientific notation
            poaf_values = pd.to_numeric(df_new['Product'], errors='coerce').fillna(0).astype(np.int64)
            
            factory_output = pd.DataFrame({
                'POAF': poaf_values,  # Now stored as integer
                'PW': '',
                'PO': df_new['PO No.'],
                'Client Ref': df_new['Client_Ref'],
                'PN': df_new['PN'],
                '固件': df_new['Firmware'],
                '固件 MPN': '',
                'Product': df_new['Product_MPN'],
                'Ordered Qty': df_new['Qty'],
                'Release Dte': datetime.today().strftime('%Y-%m-%d'),
                'Delivery date': pd.to_datetime(df_new['Deliv. Date']).dt.strftime('%Y-%m-%d'),
                'ODM': '',
                'WIZ MO number': '',
                'SO': '',
                'Client PO No.': df_new['DC PO No'],
                'Ship to code': df_new['Ship_To_Clean'],
                'Description': df_new['Product Desc.']
            })
            
            # Step 11: Generate orderhub output (UPDATED: Tax=13, Solution='N/A')
            status_text.text("Step 11/12: Generating OrderHub template...")
            progress_bar.progress(88)
            
            orderhub_output = pd.DataFrame({
                'Purchase Order No': df_new['PO No.'],
                'Remark': '',
                'Internal Part Number': df_new['Product_MPN'],
                'Quantity': df_new['Qty'],
                'Taxed Unit Price': '',
                'Untaxed Unit Price': df_new['Unit_Price'],
                'Customer Code': 'C02026',
                'Customer Name': '昕诺飞（中国）投资有限公司',
                'Seller': 'ESPDB',
                'Customer Part Number': df_new['PN'],
                'Opp Number': 'OPP-20190506-7694',
                'Tax Rate': '13%',  # Changed from 0.13 to 13
                'Currency': 'RMB',
                'Has Passed Hardware Review': 'Yes',
                'Customised/Commom': 'Customised',
                'Solution': 'N/A',  # Changed from '' to 'N/A'
                'Required Delivery Date(yyyy-MM-dd)': pd.to_datetime(df_new['Deliv. Date']).dt.strftime('%Y-%m-%d')
            })
            
            # Step 12: Generate 格式转换 export file (NEW)
            status_text.text("Step 12/12: Generating 格式转换 file...")
            progress_bar.progress(94)
            
            # Create lookup dictionary for 格式转换 from history
            # Column 19 is 合并 (UUID), Column 21 is 格式转换
            uuid_col_name = df_history.columns[19]  # 合并
            format_col_name = df_history.columns[21]  # 格式转换
            
            # Build lookup dict: UUID -> 格式转换 (as date only, no timestamp)
            format_lookup = {}
            for idx, row in df_history.iterrows():
                uuid_val = str(row[uuid_col_name]) if pd.notna(row[uuid_col_name]) else None
                format_val = row[format_col_name] if pd.notna(row[format_col_name]) else None
                if uuid_val and format_val:
                    # Convert to date only (no timestamp)
                    if hasattr(format_val, 'date'):
                        format_lookup[uuid_val] = format_val.date()
                    elif hasattr(format_val, 'strftime'):
                        format_lookup[uuid_val] = pd.to_datetime(format_val).date()
                    else:
                        format_lookup[uuid_val] = format_val
            
            # Create enhanced export (copy of original)
            df_export_enhanced = df_export.copy()
            
            # Lookup 格式转换 for each row using UUID
            def get_format_conversion(uuid):
                return format_lookup.get(str(uuid), '')
            
            df_export_enhanced['格式转换'] = df_export_enhanced['UUID'].apply(get_format_conversion)
            
            # Replace UoM column with 格式转换
            uom_col_idx = df_export_enhanced.columns.get_loc('UoM')
            df_export_enhanced.insert(uom_col_idx, '格式转换_temp', df_export_enhanced['格式转换'])
            df_export_enhanced = df_export_enhanced.drop(columns=['UoM', '格式转换'])
            df_export_enhanced = df_export_enhanced.rename(columns={'格式转换_temp': '格式转换'})
            
            # Remove helper columns for clean output (keep same as original export)
            cols_to_remove = ['Product_Base', 'Brand', 'PN', 'UUID', 'Is_New']
            df_export_enhanced = df_export_enhanced.drop(columns=[c for c in cols_to_remove if c in df_export_enhanced.columns])
            
            # Count how many 格式转换 were found
            format_found = (df_export_enhanced['格式转换'] != '').sum()
            format_missing = len(df_export_enhanced) - format_found
            
            # Complete!
            progress_bar.progress(100)
            status_text.text("✅ Processing complete!")
            
            st.success(f"🎉 Successfully processed {len(df_new)} new orders!")
            
            # Store results in session state for admin panel
            st.session_state.df_new = df_new
            st.session_state.processing_complete = True
            
            # Display results
            st.header("2️⃣ Download Your Files")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("🏭 Factory Template")
                st.dataframe(factory_output.head(3), use_container_width=True)
                
                # Convert to Excel with proper formatting for POAF
                factory_buffer = io.BytesIO()
                with pd.ExcelWriter(factory_buffer, engine='openpyxl') as writer:
                    factory_output.to_excel(writer, index=False, sheet_name='Sheet1')
                    # Format POAF column as number (no decimals)
                    worksheet = writer.sheets['Sheet1']
                    for row in range(2, len(factory_output) + 2):
                        cell = worksheet.cell(row=row, column=1)  # POAF is column A
                        cell.number_format = '0'  # Format as whole number
                factory_buffer.seek(0)
                
                output_date = datetime.now().strftime('%Y%m%d')
                st.download_button(
                    label="⬇️ Download Factory Template",
                    data=factory_buffer,
                    file_name=f"factory_output_{output_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="download_factory"
                )
            
            with col2:
                st.subheader("📋 OrderHub Template")
                st.dataframe(orderhub_output.head(3), use_container_width=True)
                
                # Convert to Excel
                orderhub_buffer = io.BytesIO()
                with pd.ExcelWriter(orderhub_buffer, engine='openpyxl') as writer:
                    orderhub_output.to_excel(writer, index=False, sheet_name='Sheet1')
                orderhub_buffer.seek(0)
                
                st.download_button(
                    label="⬇️ Download OrderHub Template",
                    data=orderhub_buffer,
                    file_name=f"orderhub_output_{output_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="download_orderhub"
                )
            
            # NEW: 格式转换 File Download
            st.subheader("📄 格式转换")
            
            col_a, col_b = st.columns([2, 1])
            with col_a:
                st.markdown(f"""
                Original export file with **Column L (UoM)** replaced by **格式转换**:
                - **{format_found}** orders have 格式转换 filled
                - **{format_missing}** orders are blank (need manual fill)
                """)
            
            with col_b:
                # Convert to Excel
                export_enhanced_buffer = io.BytesIO()
                with pd.ExcelWriter(export_enhanced_buffer, engine='openpyxl') as writer:
                    df_export_enhanced.to_excel(writer, index=False, sheet_name='Sheet1')
                export_enhanced_buffer.seek(0)
                
                st.download_button(
                    label="⬇️ Download 格式转换",
                    data=export_enhanced_buffer,
                    file_name=f"格式转换_{output_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="download_export_enhanced"
                )
            
            # Show summary
            st.header("📊 Processing Summary")
            
            summary_col1, summary_col2, summary_col3 = st.columns(3)
            
            with summary_col1:
                st.metric("Orders Processed", len(df_new))
                st.metric("Firmware Found", firmware_found)
            
            with summary_col2:
                st.metric("Matter Products", matter_count)
                st.metric("Non-Matter Products", non_matter_count)
            
            with summary_col3:
                st.metric("Unique PO Numbers", df_new['PO No.'].nunique())
                st.metric("Firmware Missing", firmware_missing)
            
            # ==================================================================
            # AUTO-UPDATE ORDER HISTORY
            # ==================================================================
            st.header("3️⃣ Update Order History (Optional)")
            
            st.info("""
            💡 **Recommended**: Add these orders to your history file to avoid processing them again next time.
            """)
            
            # Initialize session state for history generation
            if 'history_generated' not in st.session_state:
                st.session_state.history_generated = False
            if 'history_buffer' not in st.session_state:
                st.session_state.history_buffer = None
            
            col1, col2 = st.columns([2, 1])
            
            with col1:
                st.markdown("""
                This will:
                - Add the {0} new orders to your history
                - Prevent duplicate processing next time
                - Keep your history file up to date
                """.format(len(df_new)))
            
            with col2:
                if not st.session_state.history_generated:
                    generate_history = st.button(
                        "📝 Generate Updated History", 
                        use_container_width=True,
                        type="secondary",
                        key="gen_history"
                    )
                else:
                    st.success("✅ History generated!")
            
            # Generate the history file
            if generate_history or st.session_state.history_generated:
                if not st.session_state.history_generated:
                    with st.spinner("Generating updated history file..."):
                        
                        # Prepare new rows matching history structure
                        new_rows = []
                        for idx, new_row in df_new.iterrows():
                            history_row = {col: '' for col in df_history.columns}
                            
                            # Fill in the columns we have data for
                            history_row[df_history.columns[0]] = datetime.today()  # Date
                            if 'PO' in df_history.columns:
                                history_row['PO'] = new_row['PO No.']
                            if 'Product' in df_history.columns:
                                history_row['Product'] = new_row['Product']
                            if 'ClientRef' in df_history.columns:
                                history_row['ClientRef'] = new_row['PN']
                            if 'PN' in df_history.columns:
                                history_row['PN'] = new_row['PN']
                            if 'OrderedQty' in df_history.columns:
                                history_row['OrderedQty'] = new_row['Qty']
                            
                            # UUID in column T (index 19)
                            history_row[df_history.columns[19]] = new_row['UUID']
                            
                            new_rows.append(history_row)
                        
                        # Create dataframe and append
                        new_rows_df = pd.DataFrame(new_rows)
                        updated_history = pd.concat([df_history, new_rows_df], ignore_index=True)
                        
                        # Save to buffer
                        history_buffer = io.BytesIO()
                        with pd.ExcelWriter(history_buffer, engine='openpyxl') as writer:
                            updated_history.to_excel(writer, sheet_name='1.6-now', index=False)
                        history_buffer.seek(0)
                        
                        # Store in session state
                        st.session_state.history_buffer = history_buffer
                        st.session_state.history_generated = True
                        st.rerun()
                
                # Show download button
                if st.session_state.history_buffer:
                    st.success(f"✅ Updated history file ready! Added {len(df_new)} orders.")
                    
                    # Show the counts
                    col_a, col_b, col_c = st.columns(3)
                    col_a.metric("Previous Orders", len(df_history))
                    col_b.metric("New Orders Added", len(df_new))
                    col_c.metric("Total Orders", len(df_history) + len(df_new))
                    
                    output_date = datetime.now().strftime('%Y%m%d')
                    st.download_button(
                        label="⬇️ Download Updated Order History",
                        data=st.session_state.history_buffer.getvalue(),
                        file_name=f"signify_order_list_updated_{output_date}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="download_history",
                        help="Save this file and use it as your order history next time!"
                    )
                    
                    st.info("""
                    **Next Steps:**
                    1. Download this updated history file
                    2. Replace your old `signify_order_list.xlsx` with this file
                    3. Next time you process orders, upload this updated file
                    
                    This ensures orders aren't processed twice!
                    """)
            
        except Exception as e:
            progress_bar.progress(0)
            status_text.text("")
            st.error("❌ An error occurred!")
            st.exception(e)
            st.info("💡 Please check your files and try again. If the problem persists, contact Joshua.")

else:
    st.info("👆 Please upload all three files to get started")
    
    with st.expander("📖 Need help?"):
        st.markdown("""
        ### What files do I need?
        
        1. **Order History** (`signify_order_list.xlsx`)
           - Your historical orders file
           - Used to avoid processing duplicates
        
        2. **Master Reference** (`Signify_Master_Reference_COMPLETE.xlsx`)
           - Contains firmware lookup tables
           - Contains ship-to locations
           - Contains product mappings
        
        3. **Export File** (daily file from customer)
           - The file you receive from customer each day
           - Contains new orders to process
        
        ### How do I use this?
        
        1. Upload the two reference files in the sidebar (once only)
        2. Upload today's export file in the main area
        3. Click "Process Orders"
        4. Download your two output files
        5. Done! Send/upload the files as usual
        
        ### Where are my files saved?
        
        Files are **not** saved on the server. You must download them after processing.
        The app processes in real-time and provides downloads immediately.
        
        ### Is my data secure?
        
        Yes! Files are processed in memory only and deleted immediately after.
        Nothing is stored permanently on the server.
        """)

# Footer
st.markdown("---")
st.markdown("Made by Joshua 😎 | v2.0 | January 2026")
