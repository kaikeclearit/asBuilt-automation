from docx import Document

def extract_nutanix_asbuilt(nutanix_path):
    """
    Scans the vendor Nutanix As-Built document and pulls all dynamic hardware, 
    network, and configuration data into categorized lists.
    """
    doc = Document(nutanix_path)
    
    # Initialize the dictionary with exact keys matching the Jinja tags in the Clear IT Template
    nutanix_data = {
        "hosts_details": [],
        "hosts_models": [],
        "hosts_informations": [],
        "bios_bmc": [],
        "network_information": [],
        "service_name_hour": [],
        "storage_pools": [],
        "containers_list": [],
        "containers_options": [],
        "cluster_data_services": [],
        "control_virtual_machine": [],
        "licensing": [],
        "alert_monitoring": [],
        "list_directory": [],
        "prism_element": [],
        "prism_central": []
    }

    for table in doc.tables:
        if not table.rows:
            continue
            
        header = [cell.text.strip().replace('\n', ' ') for cell in table.rows[0].cells]

        # 1. Detalhes Do Host
        if "Hostname" in header and "Hypervisor Version" in header:
            for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 2 and cells[0] != "Hostname" and cells[0]:
                    nutanix_data["hosts_details"].append({
                        "hostname": cells[0], 
                        "hypervisor_version": cells[1]
                    })

        # 2. Modelos Do Host
        elif "Hostname" in header and "Model" in header and "Serial" in header:
            for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 3 and cells[0] != "Hostname" and cells[0]:
                    nutanix_data["hosts_models"].append({
                        "hostname": cells[0], 
                        "model": cells[1], 
                        "serial": cells[2]
                    })

        # 3. Informações de CPU e Memória
        elif "Hostname" in header and "CPU Cores" in header and "Memory" in header:
            for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 4 and cells[0] != "Hostname" and cells[0]:
                    nutanix_data["hosts_informations"].append({
                        "hostname": cells[0], 
                        "cpu_cores": cells[1], 
                        "cpu_threads": cells[2], 
                        "memory": cells[3]
                    })

        # 4. BIOS e BMC
        elif "Hostname" in header and "Bios Version" in header and "BMC Version" in header:
            for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 5 and cells[0] != "Hostname" and cells[0]:
                    nutanix_data["bios_bmc"].append({
                        "hostname": cells[0],
                        "bios_version": cells[1],
                        "bios_model": cells[2],
                        "bmc_version": cells[3],
                        "bmc_model": cells[4]
                    })

        # 5. Network Information
        elif "Hostname" in header and "CVM IP" in header and "Management IP" in header:
            for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 4 and cells[0] != "Hostname" and cells[0]:
                    nutanix_data["network_information"].append({
                        "hostname": cells[0], 
                        "cvm_ip": cells[1], 
                        "mgmt_ip": cells[2], 
                        "ipmi_ip": cells[3]
                    })

        # 6. NTP / DNS (Service Name Hour)
        elif "NTP Server(s)" in header and "DNS Server(s)" in header:
            for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 3 and cells[0] != "NTP Server(s)" and cells[0]:
                    nutanix_data["service_name_hour"].append({
                        "ntp_server": cells[0], 
                        "dns_server": cells[1], 
                        "global_whitelist": cells[2]
                    })
                    
        # 7. Storage Pools
        elif "Name" in header and "Storage Pool ID" in header and "Max Capacity" in header:
            for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 4 and cells[0] != "Name" and cells[0]:
                    nutanix_data["storage_pools"].append({
                        "storage_pool_name": cells[0], 
                        "storage_pool_id": cells[1], 
                        "max_capacity": cells[2],
                        "ilm_threshold": cells[3]
                    })

        # 8. Containers List
        elif "Container Name" in header and "Max Usable Capacity" in header:
            for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 5 and cells[0] != "Container Name" and cells[0]:
                    nutanix_data["containers_list"].append({
                        "hostname": cells[0], 
                        "max_usable_capacity": cells[1],
                        "total_raw_capacity": cells[2],
                        "reserved_capacity": cells[3],
                        "nfs_whitelist": cells[4]
                    })

        # 9. Container Options
        elif "Container Name" in header and "Compression Enabled" in header:
            for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 7 and cells[0] != "Container Name" and cells[0]:
                    nutanix_data["containers_options"].append({
                        "hostname": cells[0], 
                        "rf": cells[1],
                        "compression_enabled": cells[2],
                        "compression_delay": cells[3],
                        "ssd_dedup": cells[4],
                        "hdd_dedup": cells[5],
                        "erasure_coding": cells[6]
                    })

        # 10. Control Virtual Machines (CVM)
        elif "CVM" in header and "Memory" in header and "vCPU" in header:
            for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 4 and cells[0] != "CVM" and cells[0]:
                    nutanix_data["control_virtual_machine"].append({
                        "cvm_name": cells[0],
                        "memory": cells[1],
                        "vcpu": cells[2],
                        "ip_address": cells[3]
                    })

        # 11. Licensing
        elif "License" in header and "License Type" in header and "Block Serial Number" in header:
            for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 4 and cells[0] != "License":
                    nutanix_data["licensing"].append({
                        "license_name": cells[0],
                        "license_type": cells[1],
                        "block_serial_number": cells[2],
                        "expiration": cells[3]
                    })

        # 12. Alert and Monitoring (SMTP/SNMP)
        elif "Host Name" in header and "Port" in header and "Security Mode" in header:
            for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 5 and cells[0] != "Host Name" and cells[0]:
                    nutanix_data["alert_monitoring"].append({
                        "host_name": cells[0],
                        "port": cells[1],
                        "security_mode": cells[2],
                        "username": cells[3],
                        "email_address": cells[4]
                    })

        # 13. Directory List
        elif "Directory Type" in header and "Connection Type" in header and "Domain" in header:
             for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 5 and cells[1] != "Directory Type" and cells[1]:
                    nutanix_data["list_directory"].append({
                        "directory_name": cells[1] if len(cells) > 1 else "",
                        "directory_type": cells[2] if len(cells) > 2 else "",
                        "connection_type": cells[3] if len(cells) > 3 else "",
                        "directory_url": cells[4] if len(cells) > 4 else "",
                        "directory_domain": cells[5] if len(cells) > 5 else ""
                    })

        # 14. Prism Element
        elif "Cluster Name" in header and "Cluster UUID" in header and "Cluster IP" in header:
            for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 4 and cells[0] != "Cluster Name" and cells[0]:
                    nutanix_data["prism_element"].append({
                        "cluster_name": cells[0],
                        "cluster_uuid": cells[1],
                        "cluster_ip": cells[2],
                        "cluster_rf": cells[3]
                    })
                    
        # 15. Cluster Data Services (Added for safety)
        elif "Cluster Name" in header and "Data Service IP" in header:
            for row in table.rows[1:]:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 2 and cells[0] != "Cluster Name" and cells[0]:
                    nutanix_data["cluster_data_services"].append({
                        "cluster_name": cells[0],
                        "data_service_ip": cells[1]
                    })

    return nutanix_data