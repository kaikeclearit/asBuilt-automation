import streamlit as st
import os
import glob
from src.vendor_generator import run_silent_nutanix_generation
from src.extractor import extract_nutanix_asbuilt
from src.generator import generate_clearit_doc

# 1. Page Configuration
st.set_page_config(
    page_title="Clear IT As-Built Automator",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Clear IT - Document Automation")
st.markdown("Provide the **Nutanix Cluster IP** to trigger silent generation, parsing, and custom delivery document rendering.")

# Define where your vendor automation folder is located locally
VENDOR_TOOL_DIR = r"C:\Users\KaikeMaciel\Downloads\Nutanix_Service_As_Built_Win_v33"

# Ensure our working directories exist
os.makedirs("data/templates", exist_ok=True)
os.makedirs("data/generated_docs", exist_ok=True)

# 2. Network Target Input Layer with Authentication Fields
st.markdown("### 🌐 Step 1: Nutanix Environment Target")
col_ip, col_user, col_pass = st.columns([2, 1, 1])

with col_ip:
    cluster_ip = st.text_input("Cluster IP / Prism Element IP Address", placeholder="e.g., 10.10.210.10")
with col_user:
    cluster_username = st.text_input("Username", value="admin")
with col_pass:
    cluster_password = st.text_input("Password", type="password", placeholder="••••••••")

# Initialize a dictionary to carry manually captured installer specifications safely
installer_data = {}

# We reveal the physical customization layout forms once an IP target is provided
if cluster_ip:
    # --- DYNAMIC INSTALLER INTERFACE CONFIGURATION ---
    st.markdown("---")
    st.markdown("### 🛠️ Clear IT Installer Variables")
    st.caption("Configure the physical environment specifications. The form scales dynamically based on your cluster size.")

    # Number input to dynamically alter physical field row counts
    num_hosts = st.number_input("How many hosts (nodes) are in this cluster?", min_value=1, max_value=16, value=3)

    # Group A: Dynamic Physical Infrastructure Layout
    with st.expander("📌 Physical Rack & Layout Info"):
        st.markdown("##### Rack Units (RUID)")
        col1, col2 = st.columns(2)
        for i in range(1, num_hosts + 1):
            with col1:
                installer_data[f'ruid_{i}'] = st.text_input(f"Rack Unit (RUID) - Node {i}", placeholder=f"e.g., {32 - i*2}")
            with col2:
                installer_data[f'ruid_{i}_empty'] = st.text_input(f"Empty Spacer Unit below Node {i}", placeholder=f"e.g., {31 - i*2}")

    # Group B: Dynamic Physical Network Cabling Mapping
    with st.expander("🔌 Switch Port Mapping"):
        col1, col2 = st.columns(2)
        with col1:
            installer_data['switch_1_name'] = st.text_input("TOR Switch 1 Name", placeholder="e.g., CORE-DMZ-SW1")
            installer_data['switch_2_name'] = st.text_input("TOR Switch 2 Name", placeholder="e.g., CORE-DMZ-SW2")
        
        st.markdown("---")
        st.markdown("##### Node Interfaces to Switch Ports Mapping")
        for i in range(1, num_hosts + 1):
            col_a, col_b = st.columns([1, 3])
            with col_a:
                st.markdown(f" \n**Node {i} Mapping**")
            with col_b:
                installer_data[f'sw_port_{i}'] = st.text_input(f"Switch Port connected to Node {i}", key=f"sw_port_key_{i}", placeholder="e.g., TE/0/23")

    # Group C: Hardware Specifications and Contract Details
    with st.expander("📜 Hardware Specs & Support SLA"):
        col1, col2 = st.columns(2)
        with col1:
            installer_data['node_model'] = st.text_input("Nutanix Block Model", placeholder="e.g., NX-3155-G9")
            installer_data['cpu_spec'] = st.text_input("Processor Specification", placeholder="e.g., 02 x Intel Xeon-Gold 6542Y")
            installer_data['memory_spec'] = st.text_input("Memory Specification", placeholder="e.g., 8 x 64GB 5600MHz DDR5")
            installer_data['nvme_spec'] = st.text_input("NVMe Disks Spec", placeholder="e.g., 04 x 7.68 TB NVMe SSD")
            installer_data['gpu_spec'] = st.text_input("GPU Spec (If applicable)", placeholder="e.g., 03 x GPU NVIDIA L40S")
        with col2:
            installer_data['nic_spec'] = st.text_input("Network Interfaces", placeholder="e.g., 04 x 10/25GbE SFP28")
            installer_data['psu_spec'] = st.text_input("Power Supplies (PSU)", placeholder="e.g., 02 x 2000W 240V")
            installer_data['cable_spec'] = st.text_input("Power Cables", placeholder="e.g., 02 x C13/NBR")
            installer_data['support_duration'] = st.text_input("Contract / Support Duration", placeholder="e.g., 60 meses")

    # Group D: Project Scope Observations
    with st.expander("📝 Field Installation Observations"):
        installer_data['observation_1'] = st.text_area("Observation 1", placeholder="e.g., Os hosts foram taggeados com a vlan 210.")
        installer_data['observation_2'] = st.text_area("Observation 2", placeholder="e.g., Os servidores ficaram conectados em PDUs temporárias.")
        installer_data['observation_3'] = st.text_area("Observation 3", placeholder="e.g., Os cabos IPMI não foram conectados...")

st.markdown("---")

# 3. Client Metadata Inputs (Mandatory)
st.markdown("### Client Information")
client_name = st.text_input("Client Name", placeholder="e.g., Banco Ouribank")
contract_number = st.text_input("Contract Number", placeholder="e.g., 17/2023 - NE1091")


# ---------------------------------------------------------------------------
# Helper: pick the most recently created .docx from the output directory
# (ignores Word lock files that start with ~$)
# ---------------------------------------------------------------------------
def _get_latest_docx(output_dir: str) -> str:
    candidates = [
        f for f in glob.glob(os.path.join(output_dir, "*.docx"))
        if not os.path.basename(f).startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError(f"No .docx file found in output directory: {output_dir}")
    return max(candidates, key=os.path.getctime)


# 4. Pipeline Execution Button
if st.button("Generate Clear IT Document", type="primary"):

    # Validation: Ensure environmental targets and identifiers exist
    if not cluster_ip or not cluster_username or not cluster_password:
        st.error("⚠️ Please fill in the target Cluster IP, Username, and Password.")
    elif not client_name or not contract_number:
        st.error("⚠️ Please fill in both the Client Name and Contract Number.")
    else:
        with st.spinner("Processing end-to-end automation pipeline..."):
            try:
                # --- STEP A: Run Vendor Executable Layer Silently ---
                st.info("🔄 Step 1/4 — Running Nutanix binary...")
                output_dir, log_path = run_silent_nutanix_generation(
                    tool_dir=VENDOR_TOOL_DIR,
                    customer_name=client_name,
                    target_ip=cluster_ip,
                    username=cluster_username,
                    password=cluster_password
                )

                # Pick the freshest generated .docx (not a lock file)
                raw_vendor_docx = _get_latest_docx(output_dir)
                st.info(f"✅ Step 1/4 — Raw doc: `{os.path.basename(raw_vendor_docx)}`")

                # --- STEP B: Extract Automated Data From Generated File ---
                st.info("🔄 Step 2/4 — Extracting cluster data from raw document...")
                extracted_data = extract_nutanix_asbuilt(raw_vendor_docx)
                st.info("✅ Step 2/4 — Extraction complete.")

                # --- STEP C: Merge Client Meta + Manual Installer Values ---
                st.info("🔄 Step 3/4 — Merging client and installer data...")
                extracted_data["client_name"] = client_name
                extracted_data["contract_number"] = contract_number
                for key, val in installer_data.items():
                    extracted_data[key] = val
                st.info("✅ Step 3/4 — Merge complete.")

                # --- STEP D: Document Master Generation ---
                st.info("🔄 Step 4/4 — Building Clear IT delivery document...")
                template_path = os.path.join("data", "templates", "asBuilt_Template_UEPA.docx")

                if not os.path.exists(template_path):
                    st.error(f"⚠️ Template not found at: `{template_path}`")
                else:
                    safe_client_name = client_name.replace(" ", "_").replace("/", "-")
                    output_filename = f"AsBuilt_{safe_client_name}_Generated.docx"

                    final_doc_path = generate_clearit_doc(extracted_data, template_path, output_filename)
                    st.info("✅ Step 4/4 — Document built.")

                    st.success(f"✅ Document generated successfully for **{client_name}**!")

                    # Optionally show where the debug log is
                    st.caption(f"🪵 Debug log saved to: `{log_path}`")

                    # --- STEP E: Serve Final Download ---
                    with open(final_doc_path, "rb") as file:
                        st.download_button(
                            label="⬇️ Download Final Delivery Document",
                            data=file,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )

            except FileNotFoundError as e:
                st.error(f"❌ File not found: {e}")
            except TimeoutError as e:
                st.error(f"❌ Timeout: {e}")
            except Exception as e:
                st.error(f"❌ Pipeline error: {e}")
                st.exception(e)  # shows full traceback in the UI for debugging