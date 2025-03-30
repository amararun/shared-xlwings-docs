# Excel Data Analysis & ML Workflows with xlwings Lite

This repository contains example scripts demonstrating how to use xlwings Lite for data analysis and machine learning directly within Excel.

## Tips and Instructions for AI and Human Coders

### General Tips
- Use `.options(index=False)` when writing back to Excel tables to avoid unexpected index columns being added
- Format your Excel ranges as tables for best results
- Use the console and `print()` for debugging — similar to Jupyter/Colab
- xlwings lite uses a dot py file for scripts and a requirements.py and the script is put into main.py

### Critical Limitations and Workarounds
1. **Font formatting is not supported** in xlwings Lite - avoid using `.font.bold`, `.font.size`, etc. Use normal case text instead of relying on formatting.
2. **Sheet.autofit() is not supported** in xlwings Lite - you'll get a `NotImplementedError` if you try to use it.
3. **Table formatting may fail** - always wrap table formatting in try/except blocks to prevent errors from stopping execution.
4. **HTTP Request Support**:
   - Supports all common Python HTTP libraries (requests, httpx, aiohttp)
   - Native support for pyodide's pyfetch (async JavaScript fetch wrapper)
   - requests: Synchronous requests (no async/await required)
   - aiohttp/httpx: Async support available (requires async/await syntax)
   - Example:
   ```python
   # Synchronous with requests
   import requests
   response = requests.get("https://api.example.com/data")
   
   # Async with aiohttp
   import aiohttp
   async with aiohttp.ClientSession() as session:
       async with session.get("https://api.example.com/data") as response:
           data = await response.json()
   ```
5. **For FastAPI servers** that return file responses, use `await response.text()` to extract content.
6. **Pipe-delimited data** (common in FastAPI file responses) can be parsed by splitting lines with `.split("\n")` and columns with `.split("|")`.
7. **CORS is handled by FastAPI server** - make sure your API server has proper CORS settings enabled with `CORSMiddleware` to allow cross-origin requests from Excel.
8. **When working with RexDB server** that returns results as a text file, process the response as text rather than trying to parse it as JSON.
9. **Check if sheets exist** before creating them and delete existing ones if needed to avoid errors.
10. **Understanding HTTP request security from xlwings**:
    - HTTP requests from xlwings in Excel come from the origin `https://addin.xlwings.org`
    - The request headers will include:
      - `origin: 'https://addin.xlwings.org'`
      - `referer: 'https://addin.xlwings.org/'`
      - `user-agent: '...(Windows NT ...)... Microsoft Edge WebView2...'`
    - When using IP whitelisting on your API server, you must whitelist the client's IP, not xlwings.org servers
    - Requests go directly from the user's Excel/browser to your API server (not through xlwings.org)
    - Security implementations:
      - CORS settings: Allow `https://addin.xlwings.org` origin
      - IP whitelisting: Whitelist end-user IPs (the requests come directly from them)
      - For token-based auth, include Authorization header in the `pyfetch` call
    - Example FastAPI setup with security headers logging:
      ```python
      class HeadersLoggingMiddleware(BaseHTTPMiddleware):
          async def dispatch(self, request: Request, call_next):
              origin = request.headers.get("origin", "No Origin")
              user_agent = request.headers.get("user-agent", "No UA")
              logger.info(f"Request from Origin: {origin}, UA: {user_agent}")
              return await call_next(request)
              
      app.add_middleware(HeadersLoggingMiddleware)
      app.add_middleware(
          CORSMiddleware,
          allow_origins=["https://addin.xlwings.org"],
          allow_credentials=True,
          allow_methods=["*"],
          allow_headers=["*"]
      )
      ```
11. **Excel Table Formatting Best Practices:**
    - Always ensure adequate spacing between tables to prevent overlap errors
    - For model evaluation sheets, leave at least 5-7 rows gap between major sections
    - When creating multiple tables, calculate space needed including:
      - Section headers (1 row)
      - Subheaders (1 row)
      - Table data rows
      - Table headers
    - Example spacing calculation:
    ```python
    # Good practice for table placement
    start_row = previous_table_end + 5  # 5 rows gap for small tables
    start_row = previous_table_end + 7  # 7 rows gap for larger tables
    
    # Account for headers
    data_start_row = start_row + 2  # +1 for section header, +1 for subheader
    ```

12. **Table Naming and Updates:**
    - Use descriptive but unique names for tables (e.g., "TrainMetrics", "TestMetrics")
    - When updating existing tables, either:
      a) Delete and recreate the table, or
      b) Update the range value while preserving the table format
    - Example:
    ```python
    # Updating table data while preserving format
    table.range.options(index=False).value = new_df
    
    # Or recreating table
    try:
        eval_sht["A1"].options(index=False).value = df
        table_range = eval_sht["A1"].resize(len(df) + 1, len(df.columns))
        eval_sht.tables.add(table_range, name="UniqueTableName")
    except Exception as table_err:
        print(f"⚠️ Could not format as table: {table_err}")
    ```

### Best Practices for Database Connections
1. Use `await response.text()` to get raw response content before attempting to parse
2. Always include detailed error handling with `try/except` blocks
3. Print response headers, status, and content preview for debugging
4. For the RexDB server specifically:
   - Ensure URL format has a trailing slash
   - Set `db_type` parameter to match your database (e.g., "postgresql")
   - Properly encode all parameters with `urllib.parse.quote`
   - Use GET requests with query parameters
5. When parsing results:
   - Check for pipe delimiters in content
   - Convert numeric columns with `pd.to_numeric()` wrapped in try/except
   - Keep column processing simple to avoid errors

## Example Scripts

### 1. Starter Examples from xlwings documentation
Basic examples demonstrating xlwings Lite functionality:
- Hello World
- Seaborn visualization
- Custom function insertion
- Statistical operations

```python
@script
def hello_world(book: xw.Book):
    # Scripts require the @script decorator and the type-hinted
    # book argument (book: xw.Book)
    selected_range = book.selection
    selected_range.value = "Hello World!"
    selected_range.color = "#FFFF00"  # yellow


@script
def seaborn_sample(book: xw.Book):
    # Create a pandas DataFrame from a CSV on GitHub and print its info
    df = pd.read_csv(
        "https://raw.githubusercontent.com/mwaskom/seaborn-data/master/penguins.csv"
    )
    print(df.info())

    # Add a new sheet, write the DataFrame out, and format it as Table
    sheet = book.sheets.add()
    sheet["A1"].value = "The Penguin Dataset"
    sheet["A3"].options(index=False).value = df
    sheet.tables.add(sheet["A3"].resize(len(df) + 1, len(df.columns)))

    # Add a Seaborn plot as picture
    plot = sns.jointplot(
        data=df, x="flipper_length_mm", y="bill_length_mm", hue="species"
    )
    sheet.pictures.add(plot.fig, anchor=sheet["B10"])

    # Activate the new sheet
    sheet.activate()


@script
def insert_custom_functions(book: xw.Book):
    # This script inserts the custom functions below
    # so you can try them out easily
    sheet = book.sheets.add()
    sheet["A1"].value = "This sheet shows the usage of custom functions"
    sheet["A3"].value = '=HELLO("xlwings")'
    sheet["A5"].value = "=STANDARD_NORMAL(3, 4)"
    sheet["A10"].value = "=CORREL2(A5#)"
    sheet.activate()
```



### 2. XGBoost Response Model (DEF_4K.xlsx)
This script demonstrates:
- Loading data from Excel tables
- Feature preparation and encoding
- XGBoost model training
- Model evaluation with ROC curves and decile tables
- Visualization and results export back to Excel

```python
@script
def score_and_deciles(book: xw.Book):
    print("\U0001F4CC Step 1: Loading table 'DEF'...")

    sht = book.sheets.active
    table = sht.tables['DEF']
    df_orig = table.range.options(pd.DataFrame, index=False).value
    df = df_orig.copy()
    print(f"✅ Loaded table into DataFrame with shape: {df.shape}")

    # Step 2: Prepare features and target
    X = df.drop(columns=["CUSTID", "RESPONSE_TAG"])
    y = df["RESPONSE_TAG"].astype(int)
    print("🎯 Extracted features and target")

    # Step 3: One-hot encode
    X_encoded = pd.get_dummies(X, drop_first=True)
    print(f"🔢 Encoded features. Shape: {X_encoded.shape}")

    # Step 4: Split
    X_train, X_test, y_train, y_test = train_test_split(
        X_encoded, y, test_size=0.3, random_state=42
    )
    print(f"📊 Train size: {len(X_train)}, Test size: {len(X_test)}")

    # Step 5: Train XGBoost
    model = XGBClassifier(max_depth=1, n_estimators=10, use_label_encoder=False,
                          eval_metric='logloss', verbosity=0)
    model.fit(X_train, y_train)
    print("🌲 Model trained successfully.")

    # Step 6: Score train/test
    train_probs = model.predict_proba(X_train)[:, 1]
    test_probs = model.predict_proba(X_test)[:, 1]

    # Step 7: Gini
    train_gini = 2 * roc_auc_score(y_train, train_probs) - 1
    test_gini = 2 * roc_auc_score(y_test, test_probs) - 1
    print(f"📈 Train Gini: {train_gini:.4f}")
    print(f"📊 Test Gini: {test_gini:.4f}")

    # Step 8: Decile function
    def make_decile_table(probs, actuals):
        df_temp = pd.DataFrame({"prob": probs, "actual": actuals})
        df_temp["decile"] = pd.qcut(df_temp["prob"].rank(method="first", ascending=False), 10, labels=False) + 1
        grouped = df_temp.groupby("decile").agg(
            Obs=("actual", "count"),
            Min_Score=("prob", "min"),
            Max_Score=("prob", "max"),
            Avg_Score=("prob", "mean"),
            Responders=("actual", "sum")
        ).reset_index()
        grouped["Response_Rate(%)"] = round((grouped["Responders"] / grouped["Obs"]) * 100, 2)
        grouped["Cumulative_Responders"] = grouped["Responders"].cumsum()
        grouped["Cumulative_Response_%"] = round((grouped["Cumulative_Responders"] / grouped["Responders"].sum()) * 100, 2)
        return grouped

    train_decile = make_decile_table(train_probs, y_train)
    test_decile = make_decile_table(test_probs, y_test)
    print("📋 Created decile tables.")

    # Step 9: Insert deciles into new sheet
    print("📄 Preparing to insert decile tables into new sheet...")
    sheet_name = "DEF_Score_Deciles"
    existing_sheets = [s.name for s in book.sheets]

    if sheet_name in existing_sheets:
        try:
            book.sheets[sheet_name].delete()
            print(f"🧹 Existing '{sheet_name}' sheet deleted.")
        except Exception as e:
            print(f"⚠️ Could not delete existing sheet '{sheet_name}': {e}")

    new_sht = book.sheets.add(name=sheet_name, after=sht)
    new_sht["A1"].value = "Train Deciles"
    new_sht["A2"].value = train_decile

    start_row = train_decile.shape[0] + 4
    new_sht[f"A{start_row}"].value = "Test Deciles"
    new_sht[f"A{start_row+1}"].value = test_decile
    print(f"🗘️ Decile tables inserted into sheet '{sheet_name}'")

    # Step 10: Score full dataset and append as new column
    full_probs = model.predict_proba(X_encoded)[:, 1]
    df_orig["SCORE_PROBABILITY"] = full_probs
    table.range.options(index=False).value = df_orig
    print("✅ Appended SCORE_PROBABILITY to original table without changing its structure.")

    # Step 11: Create and insert graphs into Excel
    graph_sheet_name = "DEF_Score_Graphs"
    if graph_sheet_name in existing_sheets:
        try:
            book.sheets[graph_sheet_name].delete()
            print(f"🧹 Existing '{graph_sheet_name}' sheet deleted.")
        except Exception as e:
            print(f"⚠️ Could not delete existing sheet '{graph_sheet_name}': {e}")

    graph_sht = book.sheets.add(name=graph_sheet_name, after=new_sht)

    def plot_and_insert(fig, sheet, top_left_cell, name):
        try:
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, f"{name}.png")
            fig.savefig(temp_path, dpi=150)
            print(f"🖼️ Saved plot '{name}' to {temp_path}")
            anchor_cell = sheet[top_left_cell]
            sheet.pictures.add(temp_path, name=name, update=True, anchor=anchor_cell, format="png")
            print(f"✅ Inserted plot '{name}' at {top_left_cell}")
        except Exception as e:
            print(f"❌ Failed to insert plot '{name}': {e}")
        finally:
            plt.close(fig)

    # ROC Curve
    def plot_roc(y_true, y_prob, label):
        fpr, tpr, _ = roc_curve(y_true, y_prob)
        roc_auc = auc(fpr, tpr)
        plt.plot(fpr, tpr, label=f'{label} (AUC = {roc_auc:.2f})')

    fig1 = plt.figure(figsize=(6, 4))
    plot_roc(y_train, train_probs, "Train")
    plot_roc(y_test, test_probs, "Test")
    plt.plot([0, 1], [0, 1], linestyle='--', color='gray', label='Random')
    plt.title("ROC Curve")
    plt.xlabel("False Positive Rate")
    plt.ylabel("True Positive Rate")
    plt.legend()
    plt.grid(True)
    plot_and_insert(fig1, graph_sht, "A1", name="ROC_Curve")

    # Cumulative Gain Curve
    def cumulative_gain_curve(y_true, y_prob):
        df = pd.DataFrame({'actual': y_true, 'prob': y_prob})
        df = df.sort_values('prob', ascending=False).reset_index(drop=True)
        df['cumulative_responders'] = df['actual'].cumsum()
        df['cumulative_pct_responders'] = df['cumulative_responders'] / df['actual'].sum()
        df['cumulative_pct_customers'] = (df.index + 1) / len(df)
        return df['cumulative_pct_customers'], df['cumulative_pct_responders']

    train_x, train_y = cumulative_gain_curve(y_train, train_probs)
    test_x, test_y = cumulative_gain_curve(y_test, test_probs)

    fig2 = plt.figure(figsize=(6, 4))
    plt.plot(train_x, train_y, label="Train")
    plt.plot(test_x, test_y, label="Test")
    plt.plot([0, 1], [0, 1], linestyle="--", color="gray", label="Random")
    plt.title("Cumulative Gain (Decile) Curve")
    plt.xlabel("Cumulative % of Customers")
    plt.ylabel("Cumulative % of Responders")
    plt.legend()
    plt.grid(True)
    plot_and_insert(fig2, graph_sht, "A20", name="Gain_Curve")

    print(f"📊 Graphs added to sheet '{graph_sheet_name}'.")

    for sht in book.sheets:
        print(f"📄 Sheet found: {sht.name}")

    try:
        book.save()
        print("📅 Workbook saved successfully.")
    except Exception as e:
        print(f"❌ Failed to save workbook: {e}")
```

### 3. Credit Card Segment Analysis (RBICC_DEC2024.xlsx)
This script shows:
- Table formatting and data preparation
- Weighted scoring implementation
- Multicollinearity analysis
- Results visualization

```python
@script
def format_rbicc_table(book: xw.Book):
    print("Starting format_rbicc_table...")

    try:
        sht = book.sheets.active
        table = sht.tables['RBICC']
        rng = table.range
        print(f"Formatting table 'RBICC' on sheet '{sht.name}' at range: {rng.address}")

        # Get headers and all data
        headers = rng[0, :].value
        data_range = rng[1:, :rng.columns.count]
        print(f"Headers found: {headers}")

        # Define formatting rules by column name
        currency_cols = [col for col in headers if 'VALUE_AMT' in col or 'TICKET' in col]
        percent_cols = [col for col in headers if '_SHARE_' in col or 'RATIO' in col]
        round_cols = [col for col in headers if col in currency_cols + percent_cols]

        # Apply formatting column by column
        for col_idx, col_name in enumerate(headers):
            col_range = rng[1:, col_idx]  # skip header

            if col_name in currency_cols:
                col_range.number_format = '#,##0.00'  # e.g., 1,234,567.89
                print(f"Formatted '{col_name}' as currency")

            elif col_name in percent_cols:
                col_range.number_format = '0.00%'     # e.g., 68.27%
                print(f"Formatted '{col_name}' as percent")

            elif col_name in round_cols:
                col_range.number_format = '0.00'      # plain float
                print(f"Formatted '{col_name}' as float")

        # Autofit everything
        sht.autofit()
        print("Formatting complete ✅")

    except Exception as e:
        print("❌ Formatting failed:", str(e))


@script
def score_credit_card_segment(book: xw.Book):
    print("Starting score_credit_card_segment...")

    try:
        sht = book.sheets.active
        table = sht.tables['RBICC']
        rng = table.range
        print(f"Loaded table 'RBICC' from sheet: {sht.name}")

        df = table.range.options(pd.DataFrame, index=False).value
        print("Loaded table into DataFrame.")

        # Ensure correct type for scoring
        for col in df.columns:
            try:
                df[col] = df[col].astype(float)
            except:
                continue

        print("Converted numeric columns to float where possible.")

        # Scoring variables and weights
        features = {
            'TOTAL_CC_TXN_VOLUME_NOS': 20,
            'TOTAL_CC_TXN_VALUE_AMT': 20,
            'AVG_CC_TICKET_SIZE': 10,
            'POS_SHARE_OF_CC_VOLUME': 10,
            'ECOM_SHARE_OF_CC_VOLUME': 10,
            'CC_TO_DC_TXN_RATIO': 15,
            'CC_TO_DC_VALUE_RATIO': 15,
        }

        score = pd.Series(0.0, index=df.index)

        for col, weight in features.items():
            if col not in df.columns:
                print(f"Missing column for scoring: {col}")
                continue

            col_min = df[col].min()
            col_max = df[col].max()
            print(f"Normalizing {col} (min={col_min}, max={col_max})")

            # Avoid divide-by-zero
            if col_max - col_min == 0:
                normalized = 0
            else:
                normalized = (df[col] - col_min) / (col_max - col_min)

            score += normalized * weight

        df['CREDIT_CARD_SCORE'] = score.round(2)
        print("Scoring complete. Added 'CREDIT_CARD_SCORE' column.")

        # Update the table in place with new column
        table.range.value = df
        print("Updated table with score column.")

        sht.autofit()
        print("✅ Score calculation and insertion complete.")

    except Exception as e:
        print("❌ Error during scoring:", str(e))


@script
def generate_multicollinearity_matrix(book: xw.Book):
    print("Starting generate_multicollinearity_matrix...")

    try:
        sht = book.sheets.active
        table = sht.tables['RBICC']
        print(f"Loaded table 'RBICC' from sheet: {sht.name}")

        # Load data
        df = table.range.options(pd.DataFrame, index=False).value
        print("Table loaded into DataFrame.")

        # Keep only numeric columns
        numeric_df = df.select_dtypes(include='number')
        print(f"Selected {len(numeric_df.columns)} numeric columns.")

        # Calculate correlation matrix
        corr_matrix = numeric_df.corr().round(2)
        print("Correlation matrix calculated.")

        # Write to new sheet
        if 'RBICC_CorrMatrix' in [s.name for s in book.sheets]:
            book.sheets['RBICC_CorrMatrix'].delete()

        corr_sht = book.sheets.add('RBICC_CorrMatrix')
        corr_sht.range("A1").value = corr_matrix
        corr_sht.autofit()

        print("✅ Correlation matrix inserted into 'RBICC_CorrMatrix'.")

    except Exception as e:
        print("❌ Error during correlation matrix generation:", str(e))
```

### 4. Database Connection Examples (RexDB Server)
These scripts demonstrate how to connect to a database API and handle the responses:

```python
@script
async def download_file_response(book: xw.Book):
    """Access the API response as a file download."""
    print("⭐⭐⭐ STARTING download_file_response ⭐⭐⭐")
    
    try:
        # Use the *exact* URL format that worked from server logs (with trailing slash)
        api_url = "https://rexdb.hosting.tigzig.com/connect-db/"
        
        # Database connection details - matching the successful request in logs
        params = {
            "host": "YOUR_DATABASE_HOST",  # Replace with actual host from environment variables
            "database": "defaultdb", 
            "user": "avnadmin",
            "password": "YOUR_AIVEN_PASSWORD_HERE",  # Replace with actual password from environment variables
            "sqlquery": "SELECT COUNT(*) AS table_count FROM information_schema.tables WHERE table_schema = 'public';",
            "port": 13288,
            "db_type": "postgresql"
        }
        
        # Create query string with proper encoding
        query_parts = []
        for k, v in params.items():
            encoded_value = urllib.parse.quote(str(v))
            query_parts.append(f"{k}={encoded_value}")
        
        query_string = "&".join(query_parts)
        full_url = f"{api_url}?{query_string}"
        
        print(f"📤 DEBUG: Sending GET request to API...")
        
        # Make API request explicitly as a blob to properly handle file downloads
        response = await pyfetch(
            full_url,
            method="GET",
            headers={"Accept": "text/plain,application/json"},
            response_type="blob"
        )
        
        print(f"📥 DEBUG: Response status: {response.status}")
        print(f"📥 DEBUG: Response headers: {response.headers}")
        
        # Create a result sheet
        sheet_name = "File Download Result"
        if sheet_name in [s.name for s in book.sheets]:
            book.sheets[sheet_name].delete()
        
        sheet = book.sheets.add(name=sheet_name)
        
        # Simple headers without formatting
        sheet["A1"].value = "API File Response"
        
        sheet["A3"].value = "Response Status:"
        sheet["B3"].value = response.status
        
        sheet["A4"].value = "Content Type:"
        sheet["B4"].value = content_type = response.headers.get('content-type', 'unknown')
        
        sheet["A5"].value = "Content Disposition:"
        sheet["B5"].value = content_disposition = response.headers.get('content-disposition', '')
        
        # Get the raw file content as text
        text_content = await response.text()
        print(f"📥 DEBUG: Received text content: {text_content}")
        
        # Process content - format depends on the API's output
        sheet["A7"].value = "File Content (Text):"
        sheet["B7"].value = text_content
        
        # If the content is pipe-delimited (common in FastAPI's file responses)
        if "|" in text_content:
            print("📥 DEBUG: Detected pipe-delimited data, parsing as table...")
            
            # Parse pipe-delimited content into rows and columns
            lines = text_content.strip().split("\n")
            headers = [h.strip() for h in lines[0].split("|")]
            
            data_rows = []
            for line in lines[1:]:
                if line.strip():  # Skip empty lines
                    data_rows.append([cell.strip() for cell in line.split("|")])
            
            # Create a DataFrame - keeping it simple
            if data_rows:
                df = pd.DataFrame(data_rows, columns=headers)
                print(f"📊 Created DataFrame with {len(df)} rows and {len(df.columns)} columns")
                
                # Convert numeric columns
                for col in df.columns:
                    try:
                        df[col] = pd.to_numeric(df[col])
                    except:
                        pass  # Keep as string if not numeric
                
                # Display as table
                sheet["A9"].value = "Data as Table:"
                sheet["A10"].value = df
            else:
                sheet["A9"].value = "No data rows found in response"
        
        print("✅ File response successfully processed")
        
    except Exception as e:
        print(f"❌ DEBUG ERROR: {type(e).__name__}: {str(e)}")
        import traceback
        print(traceback.format_exc())
        
    print("⭐⭐⭐ ENDING download_file_response ⭐⭐⭐")

@script
async def list_all_tables(book: xw.Book):
    """List all tables in the database with column counts."""
    print("⭐⭐⭐ STARTING list_all_tables ⭐⭐⭐")
    
    try:
        # API endpoint with trailing slash
        api_url = "https://rexdb.hosting.tigzig.com/connect-db/"
        
        # Query to get all tables with column counts
        sql_query = """
        SELECT 
            table_name,
            (SELECT COUNT(*) FROM information_schema.columns c 
             WHERE c.table_schema = t.table_schema AND c.table_name = t.table_name) AS column_count
        FROM 
            information_schema.tables t
        WHERE 
            table_schema = 'public'
        ORDER BY
            table_name;
        """
        
        # Database connection details
        params = {
            "host": "rextemp-postgres-project-rextemp.h.aivencloud.com",
            "database": "defaultdb", 
            "user": "avnadmin",
            "password": "YOUR_AIVEN_PASSWORD_HERE",  # Replace with actual password from environment variables
            "sqlquery": sql_query,
            "port": 13288,
            "db_type": "postgresql"
        }
        
        # Create query string with proper encoding
        query_parts = []
        for k, v in params.items():
            encoded_value = urllib.parse.quote(str(v))
            query_parts.append(f"{k}={encoded_value}")
        
        query_string = "&".join(query_parts)
        full_url = f"{api_url}?{query_string}"
        
        print(f"📤 DEBUG: Sending GET request to API...")
        
        # Make API request with blob response type
        response = await pyfetch(
            full_url,
            method="GET",
            headers={"Accept": "text/plain,application/json"},
            response_type="blob"
        )
        
        print(f"📥 DEBUG: Response status: {response.status}")
        
        # Create a result sheet
        sheet_name = "Database Tables"
        if sheet_name in [s.name for s in book.sheets]:
            book.sheets[sheet_name].delete()
        
        sheet = book.sheets.add(name=sheet_name)
        
        # Simple header
        sheet["A1"].value = "PostgreSQL Database Tables"
        
        # Process the response
        if response.ok:
            # Get text content directly
            text_content = await response.text()
            print(f"📥 DEBUG: Response text content: {text_content}")
            
            # Process pipe-delimited content
            if "|" in text_content:
                try:
                    # Parse pipe-delimited content into rows and columns
                    lines = text_content.strip().split("\n")
                    headers = [h.strip() for h in lines[0].split("|")]
                    
                    data_rows = []
                    for line in lines[1:]:
                        if line.strip():  # Skip empty lines
                            data_rows.append([cell.strip() for cell in line.split("|")])
                    
                    # Create a DataFrame
                    if data_rows:
                        df = pd.DataFrame(data_rows, columns=headers)
                        print(f"📊 Found {len(df)} tables in database")
                        
                        # Convert numeric columns (especially the column_count)
                        for col in df.columns:
                            try:
                                df[col] = pd.to_numeric(df[col])
                            except:
                                pass  # Keep as string if not numeric
                        
                        # Display as table
                        sheet["A3"].value = df
                        
                        # Format as table - wrapped in try/except
                        try:
                            table_range = sheet["A3"].resize(len(df) + 1, len(df.columns))
                            sheet.tables.add(table_range)
                        except Exception as table_err:
                            print(f"⚠️ Could not format as table: {table_err}")
                            
                        # Add explanation for how to view table data
                        sheet["A" + str(6 + len(df))].value = "To view data for a specific table:"
                        sheet["A" + str(7 + len(df))].value = "Run the 'get_table_data' function"
                    else:
                        sheet["A3"].value = "No tables found in database"
                
                except Exception as parse_err:
                    print(f"❌ DEBUG: Table parsing error: {parse_err}")
                    sheet["A3"].value = "Error parsing table data:"
                    sheet["B3"].value = str(parse_err)
                    sheet["A4"].value = "Raw response:"
                    sheet["B4"].value = text_content
            else:
                sheet["A3"].value = "Unexpected format - raw content:"
                sheet["B3"].value = text_content
        else:
            sheet["A3"].value = "Error retrieving tables:"
            sheet["B3"].value = f"Status: {response.status}"
        
        print("✅ Table list retrieved successfully")
        
    except Exception as e:
        print(f"❌ DEBUG ERROR: {str(e)}")
        import traceback
        print(traceback.format_exc())
        
    print("⭐⭐⭐ ENDING list_all_tables SUCCESSFULLY ⭐⭐⭐") 

@script
async def get_table_data(book: xw.Book):
    """Get the data for a specific table."""
    print("⭐⭐⭐ STARTING get_table_data ⭐⭐⭐")
    
    try:
        # Get input from user for table name
        table_name = js.prompt("Enter table name to query:", "customers")
        if not table_name:
            raise Exception("No table name provided")
            
        print(f"📊 Querying data for table: {table_name}")
        
        # API endpoint with trailing slash
        api_url = "https://rexdb.hosting.tigzig.com/connect-db/"
        
        # Create SQL query for the specific table
        sql_query = f"SELECT * FROM {table_name} LIMIT 100"
        
        # Database connection details
        params = {
            "host": "rextemp-postgres-project-rextemp.h.aivencloud.com",
            "database": "defaultdb", 
            "user": "avnadmin",
            "password": "YOUR_AIVEN_PASSWORD_HERE",  # Replace with actual password from environment variables
            "sqlquery": sql_query,
            "port": 13288,
            "db_type": "postgresql"
        }
        
        # Create query string with proper encoding
        query_parts = []
        for k, v in params.items():
            encoded_value = urllib.parse.quote(str(v))
            query_parts.append(f"{k}={encoded_value}")
        
        query_string = "&".join(query_parts)
        full_url = f"{api_url}?{query_string}"
        
        print(f"📤 DEBUG: Sending GET request to API...")
        
        # Make API request for file download
        response = await pyfetch(
            full_url,
            method="GET",
            headers={"Accept": "text/plain,application/json"},
            response_type="blob"
        )
        
        print(f"📥 DEBUG: Response status: {response.status}")
        
        # Create a result sheet
        sheet_name = f"Table: {table_name}"
        if sheet_name in [s.name for s in book.sheets]:
            book.sheets[sheet_name].delete()
        
        sheet = book.sheets.add(name=sheet_name)
        
        # Simple header
        sheet["A1"].value = f"Data from table: {table_name}"
        
        # Process the response
        if response.ok:
            # Get text content directly
            text_content = await response.text()
            print(f"📥 DEBUG: Received content with length: {len(text_content)}")
            print(f"📥 DEBUG: Content preview: {text_content[:200]}...")
            
            # Process pipe-delimited content
            if "|" in text_content:
                try:
                    # Parse pipe-delimited content into rows and columns
                    lines = text_content.strip().split("\n")
                    headers = [h.strip() for h in lines[0].split("|")]
                    
                    data_rows = []
                    for line in lines[1:]:
                        if line.strip():  # Skip empty lines
                            data_rows.append([cell.strip() for cell in line.split("|")])
                    
                    # Create a DataFrame
                    if data_rows:
                        df = pd.DataFrame(data_rows, columns=headers)
                        print(f"📊 Created DataFrame with {len(df)} rows and {len(df.columns)} columns")
                        
                        # Convert numeric columns
                        for col in df.columns:
                            try:
                                df[col] = pd.to_numeric(df[col])
                            except:
                                pass  # Keep as string if not numeric
                        
                        # Display as table
                        sheet["A3"].value = df
                        
                        # Format as table if there are rows
                        if len(df) > 0:
                            try:
                                table_range = sheet["A3"].resize(len(df) + 1, len(df.columns))
                                sheet.tables.add(table_range)
                            except Exception as table_err:
                                print(f"⚠️ Could not format as table: {table_err}")
                    else:
                        sheet["A3"].value = "No data rows found in response"
                
                except Exception as parse_err:
                    print(f"❌ DEBUG: Table parsing error: {parse_err}")
                    sheet["A3"].value = "Error parsing table data:"
                    sheet["B3"].value = str(parse_err)
                    sheet["A4"].value = "Raw content:"
                    sheet["B4"].value = text_content
            else:
                # Display raw text if not pipe-delimited
                sheet["A3"].value = "Raw Response:"
                sheet["B3"].value = text_content
        else:
            error_text = await response.text()
            sheet["A3"].value = "Error:"
            sheet["B3"].value = f"Status: {response.status}"
            sheet["A4"].value = "Details:"
            sheet["B4"].value = error_text
        
        print("✅ Table data retrieved successfully")
        
    except Exception as e:
        print(f"❌ DEBUG ERROR: {str(e)}")
        import traceback
        print(traceback.format_exc())
        
    print("⭐⭐⭐ ENDING get_table_data SUCCESSFULLY ⭐⭐⭐")
```


## Setup Instructions
1. Install xlwings Add-in in Excel
2. Install required Python packages via xlwings interface for requirements.txt
   3. Open the Excel files and run the scripts using the xlwings interface


## Usage Notes
- Scripts are designed to work with specific Excel files and table structures
- Ensure data is properly formatted in Excel tables before running scripts
- Use the console output for debugging and monitoring script execution
- Results will be written to new sheets in the Excel workbook
