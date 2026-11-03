import subprocess
import re

def undeploy_all():
    print("Fetching deployments...")
    result = subprocess.run(['clasp', 'deployments'], capture_output=True, text=True, shell=True)
    
    # Parse the output
    lines = result.stdout.split('\n')
    deployment_ids = []
    
    for line in lines:
        if line.startswith('- ') and '@HEAD' not in line:
            # Example line: - AKfycbzSBgYn... @455 - Pilot App
            match = re.match(r'-\s+([a-zA-Z0-9_-]+)\s+@\d+', line)
            if match:
                deployment_ids.append(match.group(1))
                
    print(f"Found {len(deployment_ids)} deployments to remove.")
    
    for dep_id in deployment_ids:
        print(f"Undeploying {dep_id}...")
        subprocess.run(['clasp', 'undeploy', dep_id], shell=True)
        
    print("Finished undeploying unused versions.")

if __name__ == "__main__":
    undeploy_all()