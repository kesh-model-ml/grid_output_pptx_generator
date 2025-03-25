import streamlit as st
import requests
import json
import pandas as pd
import time
import io
import base64
import re
import os
from datetime import datetime

st.set_page_config(page_title="Grid to PPTX Converter", layout="wide")
st.title("Grid to PPTX Converter")

# Function to flatten the grid output JSON
# Function to flatten the grid output JSON from the specific format in strip_profile_3rows.json
def flatten_grid_json(grid_data):
    # Extract column names from the grid data
    columns = []
    if "data" in grid_data and len(grid_data["data"]) > 0:
        sheet = grid_data["data"][0]
        columns = [col["name"] for col in sheet["columns"]]
    
    # Prepare the flattened entries
    entries = []
    
    if "data" in grid_data and len(grid_data["data"]) > 0:
        sheet = grid_data["data"][0]
        
        # Create a lookup dictionary for column IDs to names
        column_id_to_name = {col["id"]: col["name"] for col in sheet["columns"]}
        
        # Create a lookup for cells by row_id and column_id
        cells_by_row_column = {}
        for cell in sheet["cells"]:
            row_id = cell["row_id"]
            column_id = cell["column_id"]
            if row_id not in cells_by_row_column:
                cells_by_row_column[row_id] = {}
            cells_by_row_column[row_id][column_id] = cell
        
        # Process each row to create an entry
        for row in sheet["rows"]:
            row_id = row["id"]
            entry = {}
            
            for column_id, column_name in column_id_to_name.items():
                if row_id in cells_by_row_column and column_id in cells_by_row_column[row_id]:
                    cell = cells_by_row_column[row_id][column_id]
                    # Clean content - remove source annotations and keep only the actual content
                    content = cell.get("content", "")
                    
                    # Remove URL source annotations
                    content = re.sub(r'<<url_source>\{.*?\}<url_source>>', '', content)
                    
                    # Handle image references
                    if "![" in content and "](attachment:" in content:
                        # Use the Airbnb logo URL as a placeholder for all images
                        content = "https://upload.wikimedia.org/wikipedia/commons/thumb/6/69/Airbnb_Logo_B%C3%A9lo.svg/1024px-Airbnb_Logo_B%C3%A9lo.svg.png?20230603231949"
                    
                    entry[column_name] = content
                else:
                    entry[column_name] = ""
            
            entries.append(entry)
    
    return {"columns": columns, "entries": entries} 

# Function to generate PPTX from flattened data with progress reporting
def generate_pptx(flattened_data, progress_placeholder):
    api_url = "https://alai-standalone-backend-proto.getalai.com/modelml/generate-presentation"
    
    # Extract column names and entries from the flattened data
    if flattened_data and "columns" in flattened_data and "entries" in flattened_data:
        columns = flattened_data["columns"]
        entries = flattened_data["entries"]
    else:
        return None, "No data to process or invalid data format"
    
    # Prepare the payload
    payload = {
        "tab": "Tab 1",
        "columns": columns,
        "entries": entries,
        "template_type": "TWO_COLUMN"  # Can be made configurable
    }
    
    headers = {
        "x-api-key": st.secrets["api_tokens"]["pptx_api_key"],
        "Accept": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "Content-Type": "application/json"
    }
    
    # Create a more dynamic progress display
    progress_placeholder.markdown("**🚀 Generating PPTX...**")
    
    # Use a spinner with status text instead of the static curl-like display
    status_container = st.empty()
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Start timer
        start_time = time.time()
        
        # Get payload size
        payload_size = len(json.dumps(payload).encode('utf-8'))
        status_container.markdown(f"📤 **Sending request:** {payload_size/1024:.1f} KB of data")
        
        # Create a session for better control
        session = requests.Session()
        
        # Show initial animation while sending request
        for i in range(10):
            # Pulse the progress bar to show activity
            progress_bar.progress(0.05 + (i * 0.02))
            status_text.markdown(f"⏳ Establishing connection... ({i+1}/10)")
            time.sleep(0.2)
            
        # Use stream=True to get response in chunks
        response = session.post(
            api_url, 
            headers=headers, 
            json=payload, 
            stream=True
        )

        # Make the API request
        status_container.markdown("📡 **API connection established**")
        progress_bar.progress(0.25)
        status_text.markdown("⏳ Sending data to presentation server...")
        
        # Show progress while receiving response
        if response.status_code == 200:
            status_container.markdown("✅ **Request accepted, downloading presentation**")
            
            # Get content length if available
            total_size = int(response.headers.get('content-length', 0))
            downloaded = 0
            chunks = []
            
            # If we don't have content length, use an indeterminate progress display
            if total_size == 0:
                for i in range(20):
                    progress_value = 0.3 + (i * 0.03)
                    progress_bar.progress(min(0.9, progress_value))
                    status_text.markdown(f"⏳ Generating presentation... ({int(progress_value*100)}%)")
                    time.sleep(0.3)
            else:
                # We have content length, show actual download progress
                for chunk in response.iter_content(chunk_size=4096):
                    if chunk:
                        chunks.append(chunk)
                        downloaded += len(chunk)
                        
                        # Calculate and show progress
                        progress_value = 0.3 + (0.7 * downloaded / total_size)
                        progress_bar.progress(min(0.95, progress_value))
                        
                        # Show download stats
                        elapsed = time.time() - start_time
                        download_speed = downloaded / elapsed if elapsed > 0 else 0
                        
                        status_text.markdown(
                            f"⏬ Downloading: {downloaded/1024:.1f} KB of {total_size/1024:.1f} KB " +
                            f"({int(downloaded/total_size*100)}%) at {download_speed/1024:.1f} KB/s"
                        )
                        
                        # Small delay to avoid too many updates
                        time.sleep(0.1)
            
            # Combine chunks to get the full response content
            content = b''.join(chunks) if chunks else response.content
            
            # Complete the progress bar
            progress_bar.progress(1.0)
            
            # Show completion message
            elapsed_time = time.time() - start_time
            status_container.markdown(f"✅ **PPTX generated successfully!**")
            status_text.markdown(
                f"⏱️ Total time: {elapsed_time:.1f} seconds | " +
                f"📊 Size: {len(content)/1024:.1f} KB"
            )
            
            return content, None
        else:
            # Request failed
            progress_bar.empty()
            status_container.markdown(f"❌ **API request failed**")
            status_text.markdown(f"Error {response.status_code}: {response.text}")
            return None, f"Error: {response.status_code} - {response.text}"
            
    except Exception as e:
        # Exception occurred
        elapsed_time = time.time() - start_time
        progress_bar.empty()
        status_container.markdown(f"❌ **Error occurred**")
        status_text.markdown(f"Exception after {elapsed_time:.1f} seconds: {str(e)}")
        return None, f"Exception: {str(e)}"

# Function to create a download link
def get_download_link(file_content, file_name):
    b64 = base64.b64encode(file_content).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="{file_name}">Download PPTX</a>'

# Function to save PPTX to a file
def save_pptx_locally(content, grid_id):
    # Use the user's Downloads folder
    downloads_dir = os.path.expanduser("~/Downloads")
    
    # Generate filename with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{downloads_dir}/grid_{grid_id}_{timestamp}.pptx"
    
    # Write file
    with open(filename, "wb") as f:
        f.write(content)
    
    return filename

# Main app flow - simplified to only use Grid URL
st.subheader("Enter Grid URL to Generate PPTX")

# Get Grid URL from user
grid_url = st.text_input("Enter Grid URL (e.g., https://app.modelml.com/grid/0195cce0-897e-79e7-b5f0-ef3f9ad09b86)")
grid_data = None

# Extract grid ID and fetch data when button is clicked
if grid_url:
    grid_id_match = re.search(r'/grid/([a-zA-Z0-9-]+)', grid_url)
    if grid_id_match:
        grid_id = grid_id_match.group(1)
        st.info(f"Grid ID: {grid_id}")
        
        # Create a button first (keep it outside the expander)
        generate_button = st.button("Generate PPTX from Grid")
        
        # Add a toggle for API details - but don't show payload until after the button is clicked
        api_details_expander = st.expander("Show API Request Details")
        with api_details_expander:
            st.code(f'''
# Grid API Request
GET https://api.modelml.com/v2/grids/grid/{grid_id}
Headers:
  accept: application/json
  X-API-KEY: [HIDDEN]
''', language="bash")
            
            # We'll update this with the payload after the button is clicked
            request_payload_placeholder = st.empty()
        
        if generate_button:
            # Create placeholders for progress updates
            fetch_progress = st.empty()
            process_progress = st.empty()
            pptx_progress = st.empty()
            
            # Step 1: Fetch Grid Data
            fetch_start_time = time.time()
            fetch_progress.text("Fetching grid data...")
            
            try:
                # Add proper headers for the API request
                headers = {
                    'accept': 'application/json',
                    'X-API-KEY': st.secrets["api_tokens"]["modelml_grid_token"]
                }
                
                # Make the API request with headers
                response = requests.get(
                    f"https://api.modelml.com/v2/grids/grid/{grid_id}",
                    headers=headers
                )
                
                fetch_elapsed = time.time() - fetch_start_time
                
                if response.status_code == 200:
                    grid_data = response.json()
                    fetch_progress.success(f"Grid data fetched successfully in {fetch_elapsed:.2f} seconds!")
                else:
                    fetch_progress.error(f"Failed to fetch grid data: {response.status_code} (after {fetch_elapsed:.2f} seconds)")
                    st.stop()
            except Exception as e:
                fetch_elapsed = time.time() - fetch_start_time
                fetch_progress.error(f"Error: {str(e)} (after {fetch_elapsed:.2f} seconds)")
                st.stop()
            
            # Step 2: Flatten the Data
            if grid_data:
                process_start_time = time.time()
                process_progress.text("Processing data...")
                
                flattened_data = flatten_grid_json(grid_data)
                process_elapsed = time.time() - process_start_time
                
                process_progress.success(f"Data processed in {process_elapsed:.2f} seconds")
                
                # Show a small preview
                st.write("Data preview:")
                if "entries" in flattened_data and flattened_data["entries"]:
                    preview_df = pd.DataFrame(flattened_data["entries"]).head(2)
                    st.dataframe(preview_df)
            
            # Create the payload for PPTX generation
            if flattened_data and "columns" in flattened_data and "entries" in flattened_data:
                payload = {
                    "tab": "Tab 1",
                    "columns": flattened_data["columns"],
                    "entries": flattened_data["entries"],
                    "template_type": "TWO_COLUMN"
                }
                
                # Update the API details expander with the complete request including payload
                with api_details_expander:
                    request_payload_placeholder.code(f'''
# Complete PPTX API Request with Payload
POST https://alai-standalone-backend-proto.getalai.com/modelml/generate-presentation
Headers:
  x-api-key: [HIDDEN]
  Accept: application/vnd.openxmlformats-officedocument.presentationml.presentation
  Content-Type: application/json

Payload:
{json.dumps(payload, indent=2)}
''', language="json")
                
                # Step 3: Generate PPTX
                pptx_content, error = generate_pptx(flattened_data, pptx_progress)
                
                if error:
                    st.error(error)
                else:
                    # Save to Downloads folder
                    local_file = save_pptx_locally(pptx_content, grid_id)
                    
                    st.success(f"PPTX generated successfully! File saved to: {local_file}")
                    
                    # Only provide a single download button
                    st.download_button(
                        label="Download PPTX",
                        data=pptx_content,
                        file_name=f"grid_{grid_id}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True  # Make the button take the full width
                    )
    else:
        st.error("Could not extract a valid Grid ID from the URL. Please check the format.")