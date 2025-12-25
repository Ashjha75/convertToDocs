import sys
import os
import re
import random
import textwrap
import traceback

# Dependency Check
try:
    from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageEnhance
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
except ImportError as e:
    print("ERROR: Missing dependencies.")
    print(f"Details: {e}")
    print("Please run: pip install Pillow reportlab")
    sys.exit(1)

# ==========================================
# CONFIGURATION
# ==========================================
OUTPUT_DIR = "output_notes"
FONT_FILENAME = "handwriting.ttf" # Expects this file in current directory

# Colors
COLOR_INK_BLUE = (11, 87, 164)   
COLOR_MARKER_RED = (198, 40, 40) 
COLOR_HIGHLIGHT = (255, 241, 118, 100) 
COLOR_PAPER = (248, 248, 245)    
COLOR_LINE_BLUE = (200, 210, 230)
COLOR_MARGIN_RED = (230, 150, 150)

# A4 Dimensions @ 300 DPI
WIDTH = 2480
HEIGHT = 3508
MARGIN_LEFT = 280
MARGIN_TOP = 150
LINE_HEIGHT_PX = 80 

# Content
RAW_PDF_CONTENT = """
--- PAGE 1 ---
Networking for DevOps

What is a Network?
When two or more computers and computing devices connected together with each other through communication channels, such as cables or wireless media and sharing some files, then it is called a Network.

A network is used to:
- Allow the connected devices to communicate with each other.
- Enable multiple users to share devices over the network, such as music and video servers, printers and scanners.

The Internet is the largest network in the world and can be called "the network of networks".

Types of Networks
There are different types of networks. But the main two are LAN and WAN.

1. LAN (Local Area Network) - interconnects computer within a limited area, such as residences, schools. e.g.: Wi-Fi, Ethernet
2. MAN (Metropolitan area network) - used in metropolitan area (cities).
3. WAN (Wide Area Network) - extends LAN over a large geographic area. e.g: optical fiber cable.
4. SONET (Synchronous Optical Network) - used in submarine.

Network Components:
1. Switch: It is a device which connects two or more computers.
2. Router: It is a device which is actually used to connect one network with another.
3. Modem:

--- PAGE 2 ---
It is also a device used for modulation and Demodulation.

4. Hub: It is just a power extension dummy device that just broadcast the signals to its connected computers.
5. NIC: It is known as Network Interface Card which is used to connect your computer with the internet. It is wireless card preinstalled on motherboard now-a-days. It has a MAC (Media Access Control) address.
6. Bridge: It is also a networking device that connects multiple LANs (local area networks) together to form a larger LAN. It reduces the broadcasting part, and it store the MAC address of the computer but now this device is also obsoleted and replaced by switch.

What is Protocol?
A network protocol is a set of rules which is set up by people that determine how a particular data is transmitted between different devices in the same network. e.g.: HTTP, TCP, IP, FTP, SMTP etc.

IP Address and its Types and Classes:
IP Address: An IP (Internet Protocol) address is a unique number assigned to each device on a network, allowing them to communicate with each other. It's like a device's "address" on the internet or local network.

Types of IP Addresses
1. IPv4: 32-bit address, written as four numbers separated by dots (e.g., 123.89.46.7).

--- PAGE 3 ---
This is a 32-bit IP address, means it contains a combo of 32 (1 and 0's). In this version of IP address there are 4 groups or Octets (8 bits), and each octet is represented by a decimal value in the address. It is easy to remember.

IPv4 Address Format (Dotted Decimal Notation)
123.89.46.72
First Octet | Second Octet | Third Octet | Fourth Octet
01111011.01011001.00101110.01001000
1 Byte = 8 Bits. 4 Bytes = 32 Bits.

Commonly used, but limited number of addresses (about 4.3 billion).

2. IPv6: 128-bit address, written in eight groups of hexadecimal numbers.
Provides a vastly larger pool of addresses, designed to replace IPv4 as it runs out.

3. Public IP: Used to identify devices on the internet. Assigned by ISPs and accessible globally.
4. Private IP: Used within private networks (like home or office networks). Not accessible from the internet; usually in ranges like 192.168.x.

--- PAGE 4 ---
10.x.x.x, or 172.16.x.x - 172.31.x.x.

5. Static IP: Manually assigned, doesn't change. Often used for servers and devices that need a consistent address.
6. Dynamic IP: Automatically assigned by a DHCP (Dynamic Host Configuration Protocol) server. Changes periodically; commonly used for home devices.

IP Address Classes (IPv4 Only)
There is an organization called IANA (Internet Assigned Numbers Authority) who divides the IP address into different classes.

IPv4 addresses are divided into five classes:
[TABLE]
Class, Range, Purpose
A, 1.0.0.0 - 126.0.0.0, Large networks (big organizations)
B, 128.0.0.0 - 191.255.0.0, Medium-sized networks
C, 192.0.0.0 - 223.255.255.0, Small networks (LANs)
D, 224.0.0.0 - 239.255.255.255, Reserved for multicasting
E, 240.0.0.0 - 255.255.255.255, Experimental/Research
[/TABLE]

--- PAGE 5 ---
Binary Representation Note:
128 64 32 16 8 4 2 1
(Powers of 2)

Key Ranges:
Class A -> 0-126
Class B -> 128-191
Class C -> 192-223

Note:
Class A addresses officially start from 1.0.0.0.
The address 0.0.0.0 is reserved.
The 127.0.0.0 to 127.255.255.255 range is reserved for LOOPBACK addresses.

What is Loopback?
Loopback address allows a device to communicate with itself.
It's often used for testing network software on the local machine.

--- PAGE 6 ---
Key Points:
- 127.0.0.1 is commonly known as "localhost."
- Any IP address in the 127.x.x.x range will loop back to the same device.

IP address - Network ID and Host ID:
There are two parts to an IP address - Network ID and Host ID.
The Network ID portion differs depending on the IP class:
- Class A: 1st octet is Network ID.
- Class B: 1st and 2nd octets are Network ID.
- Class C: 1st, 2nd, and 3rd octets are Network ID.

Direct Connection: Devices with the same Network ID can connect without a router.
Router Requirement: Devices with different Network IDs need a router.

--- PAGE 7 ---
Device Connection Scenario:
1. Device A IP: 17.0.0.1 (Class A) -> Network ID: 17
2. Device B IP: 17.0.4.2 (Class A) -> Network ID: 17
Result: Same Network ID -> Direct Connection Possible.

Router Usage:
If Device A (Net ID 17) tries to talk to Device B (Net ID 192.168), they have different Network IDs.
Result: Router Needed.

Subnetting:
Divides a network into smaller, more manageable segments.
Example: 192.168.1.0/24 can be divided into subnets like 192.168.1.0/25.

--- PAGE 8 ---
Example of Subnetting:
Given network: 192.168.1.0/24 (Class C).
/24 means 24 bits for network, 8 bits for host.
Total IPs: 256.

Dividing into two equal subnets (/25):
1. Subnet 1: 192.168.1.0/25
   Range: 192.168.1.0 to 192.168.1.127
   Usable Hosts: 126.

2. Subnet 2: 192.168.1.128/25
   Range: 192.168.1.128 to 192.168.1.255
   Usable Hosts: 126.

Benefits:
1. Improves Performance (Reduces broadcast domains).
2. Enhances Security (Segregation).
3. Efficient IP Usage.

--- PAGE 9 ---
CIDR (Classless Inter-Domain Routing):
Method for allocating IP addresses that replaces the older classful system.

Common CIDR Notations:
[TABLE]
Prefix, Netmask, Addresses, Comment
/32, 255.255.255.255, 1, Single host
/25, 255.255.255.128, 128, Class C / 2
/24, 255.255.255.0, 256, Standard Class C
/16, 255.255.0.0, 65536, Standard Class B
/8, 255.0.0.0, 16777216, Standard Class A
/0, 0.0.0.0, 4 Billion+, Entire Internet
[/TABLE]

Network Models
There are mainly two types of network models:
1. OSI Reference Model
2. TCP/IP Model

--- PAGE 10 ---
1. OSI Reference Model:
The OSI (Open Systems Interconnection) Model has 7 layers.

7. Application Layer: Interfaces directly with the user (HTTP, FTP).
6. Presentation Layer: Translates data formats, encryption (SSL/TLS).
5. Session Layer: Manages communication sessions.
4. Transport Layer: Reliable data transfer (TCP, UDP).
3. Network Layer: Routing and addressing (IP).
2. Data Link Layer: Physical addressing (MAC), Switches.
1. Physical Layer: Cables, raw bits.

--- PAGE 11 ---
Data Flow Example (Person X to Person Y):
1. App Layer: Prepare message.
2. Presentation: Encrypt.
3. Session: Establish session.
4. Transport: Break into segments (TCP).
5. Network: Add IP addresses.
6. Data Link: Add MAC addresses (Frames).
7. Physical: Convert to bits (010101) -> Wire.

--- PAGE 12 ---
(Reverse process happens at Receiver side)
Physical -> Data Link -> Network -> Transport -> Session -> Presentation -> Application.

--- PAGE 13 ---
Diagram Summary of OSI Layers:
7. Application (Data)
6. Presentation (Data)
5. Session (Data)
4. Transport (Segments)
3. Network (Packets)
2. Data Link (Frames)
1. Physical (Bits)

--- PAGE 14 ---
Protocols by Layer:
7. Application: HTTP, DNS, SMTP, FTP, SSH.
6. Presentation: SSL, TLS, JPEG, ASCII.
5. Session: NetBIOS, RPC.
4. Transport: TCP, UDP.
3. Network: IP, ICMP, IPSec, NAT.
2. Data Link: Ethernet, MAC, VLAN, ARP.
1. Physical: USB, Ethernet Cables, Bluetooth.

--- PAGE 15 ---
Common Ports:
HTTP: 80
HTTPS: 443
SSH: 22
SMTP: 25
FTP: 20, 21
DNS: 53

2. TCP/IP Model:
A simplified, real-world model with 4 layers.
1. Application Layer (Combines OSI App, Pres, Session).
2. Transport Layer.
3. Internet Layer (OSI Network).
4. Network Access Layer (Combines OSI Data Link, Physical).

--- PAGE 16 ---
(TCP/IP Diagram)
Application -> Generates Data.
Transport -> Establishes connection (TCP/UDP).
Internet -> Routes packets (IP).
Network Access -> MAC addressing / Hardware.

--- PAGE 17 ---
Ports and Protocols Details:

1. HTTP (Hypertext Transfer Protocol):
Client-server stateless protocol.
Methods: GET (Fetch), POST (Send data), PUT (Update), DELETE (Remove).

Status Codes:
200: OK (Success)
301/302: Redirects
400: Bad Request
401: Unauthorized
403: Forbidden
404: Not Found
500: Internal Server Error
502: Bad Gateway
503: Service Unavailable

--- PAGE 18 ---
Cookies:
Since HTTP is stateless, cookies are used to store session data (like login status) on the client side so you don't have to log in on every refresh.

2. SMTP (Simple Mail Transfer Protocol): Sending email.
3. POP3/IMAP: Receiving email.

4. SSH (Secure Shell):
Securely access remote computers (Port 22). Replaces Telnet.

--- PAGE 19 ---
5. TCP (Transmission Control Protocol):
Reliable, connection-oriented.
3-Way Handshake:
1. Client sends SYN.
2. Server responds SYN-ACK.
3. Client sends ACK.
-> Connection Established.

6. UDP (User Datagram Protocol):
Unreliable, connectionless, fast. No handshake. Used for streaming, DNS.

--- PAGE 20 ---
Routing:
How packets move from network to network.
Routers use a Route Table to make decisions.
Rule: "Narrowest" (most specific) match wins.
If no match found -> Default Gateway (0.0.0.0/0).

--- PAGE 21 ---
DNS (Domain Name System):
Translates human names (google.com) to IPs (142.250.x.x).

Steps:
1. Browser checks cache.
2. Asks Recursive Resolver (ISP).
3. Asks Root Server (.).
4. Asks TLD Server (.com).
5. Asks Authoritative Server (google.com).
6. IP returned.

--- PAGE 22 ---
DNS Records:
A: IPv4 Address.
AAAA: IPv6 Address.
CNAME: Alias (name to name).
MX: Mail Exchange.
NS: Name Server.
TXT: Text / Verification.

--- PAGE 23 ---
DHCP (Dynamic Host Configuration Protocol):
Automatically assigns IPs, Subnet Mask, Gateway to devices.

Network Devices:
- Router: Connects networks.
- Switch: Connects devices in LAN (uses MAC).
- Firewall: Security rules (Allow/Deny).
- Load Balancer: Distributes traffic.

--- PAGE 24 ---
Troubleshooting Tools:

1. ping: Check connectivity (uses ICMP).
   `ping google.com`

2. traceroute (tracert): Trace path/hops to destination.
   `traceroute 8.8.8.8`

3. telnet: Test specific port connectivity.
   `telnet google.com 80`

4. curl: Transfer data / test HTTP.
   `curl -I http://example.com` (Check headers)

5. dig: DNS lookup details.
   `dig google.com`

--- PAGE 25 ---
6. netstat: Network statistics / open ports.
   `netstat -tuln`

7. nmap: Network mapper / security scanner.
   `nmap -sP 192.168.1.0/24`

8. ssh: Remote login.
   `ssh user@host`

9. scp: Secure copy files.
   `scp file.txt user@host:/path`

--- END ---
"""

# ==========================================
# CLASS: PAGE RENDERER
# ==========================================

class NotebookPageGenerator:
    def __init__(self, output_dir, font_path):
        self.output_dir = output_dir
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            
        self.base_font_size = 55
        self.font = self.load_font(font_path, self.base_font_size)
        self.header_font = self.load_font(font_path, int(self.base_font_size * 1.3))
        self.code_font = self.load_font(font_path, int(self.base_font_size * 0.9))

    def load_font(self, path, size):
        """Robust font loading with system fallback."""
        try:
            return ImageFont.truetype(path, size)
        except (OSError, IOError):
            # Try common system fonts if custom font is missing
            try:
                # Windows/Linux common
                return ImageFont.truetype("arial.ttf", size)
            except (OSError, IOError):
                try:
                    # Linux common
                    return ImageFont.truetype("DejaVuSans.ttf", size)
                except (OSError, IOError):
                    print(f"WARNING: No fonts found. Using default bitmap (will be small/ugly).")
                    return ImageFont.load_default()

    def get_text_bbox(self, draw, x, y, text, font):
        """Cross-version compatibility for text measurement."""
        try:
            # Pillow >= 8.0.0
            return draw.textbbox((x, y), text, font=font)
        except AttributeError:
            # Pillow < 8.0.0 fallback
            w, h = draw.textsize(text, font=font)
            return (x, y, x + w, y + h)

    def create_blank_page(self):
        img = Image.new('RGB', (WIDTH, HEIGHT), COLOR_PAPER)
        draw = ImageDraw.Draw(img, 'RGBA')

        # Ruled lines
        for y in range(MARGIN_TOP, HEIGHT, LINE_HEIGHT_PX):
            draw.line([(0, y), (WIDTH, y)], fill=COLOR_LINE_BLUE, width=2)

        # Margin line
        draw.line([(MARGIN_LEFT, 0), (MARGIN_LEFT, HEIGHT)], fill=COLOR_MARGIN_RED, width=3)
        return img, draw

    def draw_highlight(self, draw, x, y, text, font):
        bbox = self.get_text_bbox(draw, x, y, text, font)
        
        overlay = Image.new('RGBA', (WIDTH, HEIGHT), (0,0,0,0))
        d = ImageDraw.Draw(overlay)
        
        padding = 5
        # Ensure bbox has valid coordinates
        if bbox:
            d.rectangle(
                [bbox[0]-padding, bbox[1]+5, bbox[2]+padding, bbox[3]-5], 
                fill=COLOR_HIGHLIGHT
            )
        return overlay

    def render_page(self, page_num, content_lines):
        print(f"Rendering Page {page_num}...")
        img, draw = self.create_blank_page()
        
        cursor_y = MARGIN_TOP + 15
        cursor_x = MARGIN_LEFT + 40
        
        draw.text((WIDTH - 200, HEIGHT - 150), f"p. {page_num}", fill=COLOR_INK_BLUE, font=self.font)

        highlight_layer = Image.new('RGBA', (WIDTH, HEIGHT), (0,0,0,0))

        for line in content_lines:
            if not line.strip():
                cursor_y += LINE_HEIGHT_PX
                continue
            
            is_header = False
            current_font = self.font
            current_color = COLOR_INK_BLUE
            
            if line.isupper() or line.endswith(":") or line.startswith("---"):
                is_header = True
                current_font = self.header_font
                current_color = COLOR_MARKER_RED
            
            # Simple Table Handling
            if "," in line and line.count(",") > 1 and len(line) < 100 and not line.strip().startswith("["):
                parts = line.split(',')
                offset = 0
                for part in parts:
                    draw.text((cursor_x + offset, cursor_y), part.strip(), fill=COLOR_INK_BLUE, font=self.code_font)
                    offset += 500
                cursor_y += LINE_HEIGHT_PX
                continue
            
            # Remove table markers if present
            if "[TABLE]" in line or "[/TABLE]" in line:
                continue

            keywords = ["IP", "TCP", "UDP", "OSI", "HTTP", "DNS", "DevOps", "Network"]
            
            max_chars = 60 if not is_header else 40
            wrapped = textwrap.wrap(line, width=max_chars)
            
            for subline in wrapped:
                jitter_y = random.randint(-2, 2)
                
                for kw in keywords:
                    if kw in subline and not is_header:
                        hl_layer = self.draw_highlight(draw, cursor_x, cursor_y + jitter_y, subline, current_font)
                        highlight_layer = Image.alpha_composite(highlight_layer, hl_layer)
                
                draw.text(
                    (cursor_x, cursor_y + jitter_y), 
                    subline, 
                    fill=current_color, 
                    font=current_font
                )
                cursor_y += LINE_HEIGHT_PX * (1.2 if is_header else 1.0)
                
            cursor_y += 10

        img = img.convert("RGBA")
        img = Image.alpha_composite(img, highlight_layer)
        img = img.convert("RGB")
        img = img.filter(ImageFilter.GaussianBlur(0.5))
        
        filename = os.path.join(self.output_dir, f"page_{page_num:02d}.png")
        img.save(filename, "PNG")
        return filename

    def generate_pdf(self, image_files, output_pdf_name):
        print("Assembling PDF...")
        pdf_path = os.path.join(self.output_dir, output_pdf_name)
        c = canvas.Canvas(pdf_path, pagesize=A4)
        for img_path in image_files:
            a4_w, a4_h = A4
            c.drawImage(img_path, 0, 0, width=a4_w, height=a4_h)
            c.showPage()
        c.save()
        print(f"Success! PDF generated at: {pdf_path}")

# ==========================================
# MAIN EXECUTION
# ==========================================

def parse_and_generate():
    print("Initializing Generator...")
    generator = NotebookPageGenerator(OUTPUT_DIR, FONT_FILENAME)
    
    raw_pages = re.split(r'--- PAGE \d+ ---', RAW_PDF_CONTENT)
    if not raw_pages[0].strip():
        raw_pages.pop(0)
        
    generated_images = []
    
    try:
        for i, page_text in enumerate(raw_pages):
            page_num = i + 1
            lines = page_text.strip().split('\n')
            img_file = generator.render_page(page_num, lines)
            generated_images.append(img_file)
            
        generator.generate_pdf(generated_images, "Handwritten_Notes_Scanned.pdf")
        
    except Exception:
        print("\nCRITICAL ERROR DURING GENERATION:")
        traceback.print_exc()

if __name__ == "__main__":
    parse_and_generate()