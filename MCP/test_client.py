import subprocess
import json
import time

def run_test():
    print("Starting MCP server subprocess...")
    # Start the server
    p = subprocess.Popen(
        ["python", "server.py"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        bufsize=1
    )
    
    time.sleep(1) # Allow server to start
    
    # 1. Send initialize request
    init_request = {
        "jsonrpc": "2.0",
        "method": "initialize",
        "id": 1,
        "params": {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {
                "name": "test-client",
                "version": "1.0"
            }
        }
    }
    
    print("\n---> Sending 'initialize' request...")
    p.stdin.write(json.dumps(init_request) + "\n")
    p.stdin.flush()
    
    # Read response
    line = p.stdout.readline()
    print("<--- Received response:")
    print(json.dumps(json.loads(line), indent=2, ensure_ascii=False))
    
    # Send initialized notification
    init_notif = {
        "jsonrpc": "2.0",
        "method": "notifications/initialized"
    }
    print("\n---> Sending 'notifications/initialized'...")
    p.stdin.write(json.dumps(init_notif) + "\n")
    p.stdin.flush()
    
    # 2. Send tools/list request
    list_request = {
        "jsonrpc": "2.0",
        "method": "tools/list",
        "id": 2
    }
    print("\n---> Sending 'tools/list' request...")
    p.stdin.write(json.dumps(list_request) + "\n")
    p.stdin.flush()
    
    line = p.stdout.readline()
    print("<--- Received response:")
    print(json.dumps(json.loads(line), indent=2, ensure_ascii=False))
    
    # 3. Call get_contacts_count tool
    call_request = {
        "jsonrpc": "2.0",
        "method": "tools/call",
        "id": 3,
        "params": {
            "name": "get_contacts_count",
            "arguments": {}
        }
    }
    print("\n---> Calling 'get_contacts_count' tool...")
    p.stdin.write(json.dumps(call_request) + "\n")
    p.stdin.flush()
    
    line = p.stdout.readline()
    print("<--- Received response:")
    print(json.dumps(json.loads(line), indent=2, ensure_ascii=False))
    
    # Terminate process
    print("\nTerminating server...")
    p.terminate()
    try:
        p.wait(timeout=2)
    except subprocess.TimeoutExpired:
        p.kill()
    print("Done.")

if __name__ == "__main__":
    run_test()
