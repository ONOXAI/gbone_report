import xml.etree.ElementTree as ET
import pandas as pd

def map_severity(score):
    try:
        score = float(score)
        if score >= 9.0:
            return "Critical"
        elif score >= 6.0:
            return "High"
        elif score >= 3.0:
            return "Medium"
        elif score >= 1.0:
            return "Low"
        else:
            return "None"
    except:
        return "None"

def parse_xml(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()

    metadata = {
        "config_id": root.attrib.get("config_id"),
        "owner": root.findtext(".//owner/name"),
        "created": root.findtext(".//creation_time"),
        "modified": root.findtext(".//modification_time"),
        "scan_status": root.findtext(".//scan_run_status"),
        "scan_target": root.findtext(".//task/name"),
        "hosts_count": root.findtext(".//hosts/count"),
        "vulns_count": root.findtext(".//vulns/count"),
        "closed_cves_count": root.findtext(".//closed_cves/count"),
        "os_count": root.findtext(".//os/count"),
        "apps_count": root.findtext(".//apps/count"),
        "ssl_certs_count": root.findtext(".//ssl_certs/count"),
    }

    data = []
    for result in root.findall(".//result"):
        vuln_name = result.findtext("name")
        host = result.findtext("host")
        port = result.findtext("port")
        threat = result.findtext("threat")
        cvss_score = result.findtext("severity")
        severity = map_severity(cvss_score)
        description = result.findtext("description")
        solution_node = result.find(".//solution")
        solution = solution_node.text.strip() if solution_node is not None else ""

        cves = [ref.attrib.get("id") for ref in result.findall(".//ref[@type='cve']")]
        cve_list = ", ".join(cves)

        data.append({
            "host": host,
            "port": port,
            "vuln_name": vuln_name,
            "threat": threat,
            "cvss_score": cvss_score,
            "severity": severity,
            "cves": cve_list,
            "description": description,
            "solution": solution
        })

    df = pd.DataFrame(data)
    return df, metadata
