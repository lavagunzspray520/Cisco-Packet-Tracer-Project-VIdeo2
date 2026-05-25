#!/usr/bin/env python3
"""Generate Cisco Packet Tracer Webinar Script (Word) and Presentation (PowerPoint)."""
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
from pptx.util import Inches, Pt as PptPt, Emu
from pptx.dml.color import RGBColor as PptRGB
from pptx.enum.text import PP_ALIGN

# Ensure output directory exists
os.makedirs("webinar", exist_ok=True)

# Define slide data: each entry is a dict with all info needed
SLIDES = [
    {
        "num": 1,
        "title": "Welcome to Our Cisco Packet Tracer Webinar!",
        "script": "Hello everyone! Welcome to our webinar today. We are so happy you are here. Today, we will learn about a very cool tool called Cisco Packet Tracer. It helps us build pretend computer networks on our computer. Think of it like a video game, but for networking! By the end of this webinar, you will know how to use it. Let us get started!",
        "music": "Upbeat instrumental - 'Happy Day' by Bensound",
        "animation": "Fade in with zoom effect on title",
        "emotion": "Excited and energetic - like greeting old friends",
        "bullets": ["Welcome!", "Topic: Cisco Packet Tracer", "Duration: 1+ hour", "Interactive session"],
        "bg": "D6EAF8"
    },
    {
        "num": 2,
        "title": "What Will We Learn Today?",
        "script": "Here is what we will talk about today. First, we will learn what Cisco Packet Tracer is. Then, we will see what it can do. We will learn how to download it. We will look at real examples from the Philippines. And we will also talk about common mistakes people make. Sounds fun, right?",
        "music": "Gentle acoustic guitar - 'New Beginning' by Bensound",
        "animation": "Slide in from left, bullet points appear one by one",
        "emotion": "Warm and inviting - like a teacher starting class",
        "bullets": ["What is Cisco Packet Tracer?", "Key Features", "How to Download", "Philippine Examples", "Common Mistakes", "Tips and Tricks"],
        "bg": "D5F5E3"
    },
    {
        "num": 3,
        "title": "What is Cisco Packet Tracer?",
        "script": "So, what is Cisco Packet Tracer? It is a free program made by Cisco. Cisco is a very big company that makes network equipment, like routers and switches. Packet Tracer lets you build a pretend network on your computer. You do not need real equipment. You can drag and drop devices like computers, routers, and switches. Then you connect them with cables. It is like building with digital Lego blocks!",
        "music": "Curious melody - 'Little Idea' by Bensound",
        "animation": "Zoom in on center, text fades in",
        "emotion": "Curious and explanatory - like showing a new toy",
        "bullets": ["Free program by Cisco", "Build pretend networks", "No real equipment needed", "Drag and drop devices", "Like digital Lego blocks!"],
        "bg": "FDEBD0"
    },
    {
        "num": 4,
        "title": "What is Cisco Packet Tracer? (continued)",
        "script": "Let me explain more. Packet Tracer is a simulator. That means it pretends to be a real network. You can see how data moves from one computer to another. You can watch packets travel through cables and routers. It is used by students all over the world to learn networking. And the best part? It is completely free if you sign up for Cisco NetAcad!",
        "music": "Curious melody continues - 'Little Idea' by Bensound",
        "animation": "Crossfade from previous slide",
        "emotion": "Informative and encouraging - building excitement",
        "bullets": ["It is a simulator (pretend network)", "Watch data move between computers", "See packets travel through cables", "Used by students worldwide", "Completely FREE with NetAcad account"],
        "bg": "FFFFFF"
    },
    {
        "num": 5,
        "title": "Who Made It and Why?",
        "script": "Cisco made Packet Tracer for students. They wanted students to practice networking without buying expensive equipment. A real Cisco router can cost thousands of pesos! But with Packet Tracer, you can practice for free. It is part of the Cisco Networking Academy program. Many schools in the Philippines use it for their IT classes.",
        "music": "Inspiring background - 'Evolution' by Bensound",
        "animation": "Slide in from right with company logo placeholder",
        "emotion": "Appreciative and informative - showing value",
        "bullets": ["Made by Cisco for students", "Practice without expensive equipment", "Real routers cost thousands of pesos", "Part of Cisco Networking Academy", "Many Philippine schools use it"],
        "bg": "D6EAF8"
    },
    {
        "num": 6,
        "title": "Key Features - Build Networks Easily",
        "script": "Let us talk about the cool things Packet Tracer can do. First, you can build networks very easily. Just drag a router from the menu and drop it on the screen. Then drag a computer. Connect them with a cable. Done! You just built a tiny network. You can add as many devices as you want. Routers, switches, computers, servers, phones, even wireless access points!",
        "music": "Upbeat electronic - 'Jazzy Frenchy' by Bensound",
        "animation": "Pop-in effect for each bullet point",
        "emotion": "Enthusiastic - showing off cool features",
        "bullets": ["Drag and drop to build networks", "Routers, switches, computers", "Servers and phones too", "Wireless access points", "Connect with virtual cables", "Add as many devices as you want"],
        "bg": "D5F5E3"
    },
    {
        "num": 7,
        "title": "Key Features - Simulation Mode",
        "script": "The second amazing feature is Simulation Mode. In real life, you cannot see data moving through cables. But in Packet Tracer, you can! You click on Simulation mode, and then you can watch little envelopes move from one device to another. Each envelope is a packet of data. You can click on the envelope and see what is inside. It is like having X-ray vision for networks!",
        "music": "Wonder and discovery - 'Creative Minds' by Bensound",
        "animation": "Slide transition with moving packet animation placeholder",
        "emotion": "Amazed and wonder-filled - like showing magic",
        "bullets": ["Watch data move through cables!", "Little envelopes = packets", "Click to see inside packets", "Like X-ray vision for networks", "Pause and step through time", "Great for understanding how networks work"],
        "bg": "FDEBD0"
    },
    {
        "num": 8,
        "title": "Key Features - Practice Cisco Commands",
        "script": "The third feature is practicing Cisco commands. Real Cisco routers use a special language called IOS commands. You type commands to tell the router what to do. In Packet Tracer, you can practice these commands without a real router. Click on any router or switch, go to the CLI tab, and start typing. It feels just like the real thing! Well, almost. Some commands are missing, but most important ones work.",
        "music": "Focused study music - 'Slow Motion' by Bensound",
        "animation": "Typewriter effect on bullet points",
        "emotion": "Focused and practical - like a helpful tutor",
        "bullets": ["Practice Cisco IOS commands", "Click device then CLI tab", "Type commands like on real routers", "Most important commands work", "Some commands are missing (that is okay!)", "Great for exam preparation"],
        "bg": "FFFFFF"
    },
    {
        "num": 9,
        "title": "How to Download Cisco Packet Tracer",
        "script": "Now let us learn how to get Packet Tracer on your computer. First, you need to go to the Cisco Networking Academy website. The address is netacad.com. You need to create a free account. It asks for your name and email. After you sign up, look for the Resources section. Then find Packet Tracer. Click download. Choose the version for your computer, Windows or Mac or Linux. It is a big file, about 200 to 300 megabytes. So make sure you have good internet!",
        "music": "Step-by-step tutorial music - 'Acoustic Breeze' by Bensound",
        "animation": "Numbered steps appear one by one from top",
        "emotion": "Helpful and patient - guiding step by step",
        "bullets": ["Go to netacad.com", "Create a FREE account", "Enter name and email", "Find Resources section", "Click Download Packet Tracer", "Choose: Windows / Mac / Linux", "File size: 200-300 MB"],
        "bg": "D6EAF8"
    },
    {
        "num": 10,
        "title": "Installing Packet Tracer",
        "script": "After the download is done, double-click the file to install it. On Windows, just click Next, Next, Next, and Finish. It is very easy. On Mac, drag it to your Applications folder. When you open it for the first time, it will ask you to log in with your NetAcad account. Type your email and password. And that is it! You are ready to start building networks. The whole process takes about 5 to 10 minutes.",
        "music": "Calm progress music - 'Tomorrow' by Bensound",
        "animation": "Progress bar animation placeholder",
        "emotion": "Reassuring and simple - anyone can do this",
        "bullets": ["Double-click downloaded file", "Windows: Next, Next, Finish", "Mac: Drag to Applications", "Log in with NetAcad account", "Takes 5-10 minutes total", "You are ready to go!"],
        "bg": "D5F5E3"
    },
    {
        "num": 11,
        "title": "The Workspace - What You See When You Open It",
        "script": "When you open Packet Tracer, here is what you see. At the top, there is a menu bar with File, Edit, and other options. Below that is a toolbar with buttons for common actions. The big white area in the middle is your workspace. This is where you build your network. At the bottom, you see device categories. You pick devices from there. On the right side, there might be some extra panels. Do not worry about all the buttons. We will learn the important ones.",
        "music": "Exploration music - 'Clear Day' by Bensound",
        "animation": "Pan across interface screenshot placeholder",
        "emotion": "Patient and guiding - like a tour guide",
        "bullets": ["Menu bar at the top", "Toolbar with quick buttons", "Big white workspace in the middle", "Device categories at the bottom", "Do not worry about all buttons!", "We will learn the important ones"],
        "bg": "FDEBD0"
    },
    {
        "num": 12,
        "title": "The Workspace - Device Categories",
        "script": "Let me show you the device categories at the bottom. You will see icons for different types of devices. There are routers, which are like traffic police for data. There are switches, which connect many computers together. There are end devices like computers and laptops. There are cables to connect everything. And there are even wireless devices and IoT things. Each category has many choices inside it.",
        "music": "Exploration music continues - 'Clear Day' by Bensound",
        "animation": "Highlight each category with zoom effect",
        "emotion": "Organized and clear - categorizing things neatly",
        "bullets": ["Routers = traffic police for data", "Switches = connect many computers", "End devices = computers, laptops, phones", "Connections = cables and wires", "Wireless devices", "IoT (Internet of Things) devices"],
        "bg": "FFFFFF"
    },
    {
        "num": 13,
        "title": "The Workspace - Realtime vs Simulation Mode",
        "script": "There are two modes in Packet Tracer. Realtime mode is the default. Everything happens at normal speed, just like a real network. You cannot see individual packets moving. Simulation mode is special. It slows everything down. You can see each packet move step by step. There is a button at the bottom right to switch between them. Use Realtime mode when building. Use Simulation mode when you want to understand what is happening inside the network.",
        "music": "Thoughtful music - 'Once Again' by Bensound",
        "animation": "Split screen comparison placeholder",
        "emotion": "Clear and comparative - helping understand the difference",
        "bullets": ["Realtime Mode = normal speed", "Simulation Mode = slow motion", "Button at bottom right to switch", "Build in Realtime mode", "Learn in Simulation mode", "Simulation shows packets moving step by step"],
        "bg": "D6EAF8"
    },
    {
        "num": 14,
        "title": "Philippine Scenario: School Computer Lab in Manila",
        "script": "Let us look at real examples from the Philippines! Imagine you are setting up a computer lab in a public school in Manila. The school has 30 computers for students. They need to share one internet connection. How do we do this? In Packet Tracer, we can build this network. We put one router connected to the internet. Then a switch connects all 30 computers. We assign IP addresses to each computer. Now all students can browse the internet and do research!",
        "music": "Filipino-inspired gentle music - 'Maligaya' style acoustic",
        "animation": "Build-up animation showing network growing",
        "emotion": "Proud and relatable - showing Filipino context",
        "bullets": ["Public school in Manila", "30 computers for students", "Share one internet connection", "1 Router + 1 Switch + 30 PCs", "Assign IP addresses to each PC", "Everyone can browse and learn!"],
        "bg": "D5F5E3"
    },
    {
        "num": 15,
        "title": "Philippine Scenario: Barangay Health Center",
        "script": "Next example! A barangay health center in a small town needs internet. The health workers need to send patient reports to the city hospital. They have 5 computers and one printer. In Packet Tracer, we connect the 5 computers and the printer to a switch. We add a router for internet. Now the health workers can email reports. They can also print patient records. This is how technology helps our communities!",
        "music": "Hopeful community music - 'Pagsisikap' style",
        "animation": "Community-themed transition with map placeholder",
        "emotion": "Caring and community-focused - showing real impact",
        "bullets": ["Barangay health center", "5 computers + 1 printer", "Send reports to city hospital", "Switch connects all devices", "Router provides internet", "Technology helps communities!"],
        "bg": "FDEBD0"
    },
    {
        "num": 16,
        "title": "Philippine Scenario: Sari-Sari Store POS Network",
        "script": "Here is a fun one! A sari-sari store owner wants to modernize. They want a point-of-sale system. That means a computer to track what they sell. Maybe they also want WiFi for customers. In Packet Tracer, we set up a small network. One computer for the cash register. A wireless access point for customer WiFi. A router to connect to the internet. Even a small business can have a proper network! This is how small Filipino businesses can grow.",
        "music": "Cheerful entrepreneurial music - 'Negosyo Beat'",
        "animation": "Store-themed graphics placeholder with network overlay",
        "emotion": "Entrepreneurial and encouraging - showing possibilities",
        "bullets": ["Sari-sari store modernization", "Point-of-sale computer", "WiFi for customers", "Wireless access point", "Router for internet", "Small businesses can grow with tech!"],
        "bg": "FFFFFF"
    },
    {
        "num": 17,
        "title": "Philippine Scenario: University Campus in Cebu",
        "script": "Now let us think bigger! A university in Cebu has multiple buildings. Each building has its own network. The library, the science building, the admin office, and the dormitory. They all need to connect to each other and to the internet. In Packet Tracer, we use multiple switches and routers. We create VLANs to separate student traffic from admin traffic. This is a more complex network, but Packet Tracer can handle it!",
        "music": "Academic achievement music - 'Graduation Day' style",
        "animation": "Campus map with network connections appearing",
        "emotion": "Ambitious and academic - thinking bigger",
        "bullets": ["University with multiple buildings", "Library, Science, Admin, Dormitory", "Each building has its own network", "Multiple switches and routers", "VLANs separate different traffic", "Complex but Packet Tracer handles it!"],
        "bg": "D6EAF8"
    },
    {
        "num": 18,
        "title": "Philippine Scenario: Government Office in Davao",
        "script": "Government offices in Davao also need networks. They handle sensitive information like citizen records. Security is very important! In Packet Tracer, we can set up a secure network. We use firewalls to block bad traffic. We create separate networks for public services and private records. We add passwords to all routers and switches. This teaches us about network security, which is very important for government work.",
        "music": "Professional and secure - 'Official Business' style",
        "animation": "Security shield animation placeholder",
        "emotion": "Serious but accessible - showing importance of security",
        "bullets": ["Government office in Davao", "Handles sensitive citizen records", "Security is very important!", "Firewalls block bad traffic", "Separate public and private networks", "Passwords on all devices"],
        "bg": "D5F5E3"
    },
    {
        "num": 19,
        "title": "Philippine Scenario: Internet Cafe (Pisonet) Network",
        "script": "Who has been to a pisonet? Many Filipinos use internet cafes or pisonets for gaming and school work. Setting up a pisonet network is a great exercise! In Packet Tracer, we place 10 to 20 computers. We connect them all to a switch. We add a router with internet. We might add a server for managing time and payments. This is a real business that many Filipinos run. Learning to set up this network is very practical!",
        "music": "Gaming and fun music - 'Pixel Adventure' style",
        "animation": "Gaming-themed transition with network building up",
        "emotion": "Fun and relatable - connecting to everyday Filipino life",
        "bullets": ["Pisonet / Internet cafe setup", "10-20 computers for customers", "All connected to a switch", "Router for internet connection", "Server for time management", "Real Filipino business!"],
        "bg": "FDEBD0"
    },
    {
        "num": 20,
        "title": "Step-by-Step Tutorial: Adding Devices",
        "script": "Now let us do a simple tutorial together! Step one: Open Packet Tracer. Step two: Look at the bottom of the screen. You see device categories. Step three: Click on 'End Devices'. Step four: Find the PC icon. Step five: Click on it and then click on the workspace to place it. Do it again to place a second PC. Now you have two computers on your workspace! Easy, right?",
        "music": "Tutorial step music - 'Study Time' by Bensound",
        "animation": "Step-by-step numbered highlights",
        "emotion": "Patient teacher - going very slowly and clearly",
        "bullets": ["Step 1: Open Packet Tracer", "Step 2: Look at bottom of screen", "Step 3: Click 'End Devices'", "Step 4: Find the PC icon", "Step 5: Click PC, then click workspace", "Place two PCs on the workspace"],
        "bg": "FFFFFF"
    },
    {
        "num": 21,
        "title": "Step-by-Step Tutorial: Connecting Devices",
        "script": "Now we have two computers. But they are not connected! Let us fix that. Go to the connections category at the bottom. It looks like a lightning bolt. Click on it. Choose 'Copper Cross-Over Cable'. Why cross-over? Because we are connecting two similar devices, PC to PC. Now click on the first PC, choose FastEthernet0. Then click on the second PC, choose FastEthernet0. You see a cable connecting them! The dots will turn green when the connection is working.",
        "music": "Tutorial step music continues - 'Study Time'",
        "animation": "Cable drawing animation placeholder",
        "emotion": "Encouraging - celebrating small wins",
        "bullets": ["Go to Connections category (lightning bolt)", "Choose Copper Cross-Over Cable", "Cross-over = similar devices", "Click first PC > FastEthernet0", "Click second PC > FastEthernet0", "Green dots = connection works!"],
        "bg": "D6EAF8"
    },
    {
        "num": 22,
        "title": "Step-by-Step Tutorial: Setting IP Addresses",
        "script": "The PCs are connected, but they still cannot talk. Why? Because they do not have addresses! Think of it like houses without house numbers. The mailman cannot deliver mail. We need to give each PC an IP address. Click on the first PC. Go to the Desktop tab. Click IP Configuration. Type: 192.168.1.1 for the IP address. Subnet mask: 255.255.255.0. Do the same for the second PC but use 192.168.1.2. Now they can talk!",
        "music": "Achievement unlocked music - 'Completion' style",
        "animation": "IP address typing animation placeholder",
        "emotion": "Methodical and clear - each step matters",
        "bullets": ["PCs need IP addresses (like house numbers)", "Click PC > Desktop tab > IP Configuration", "PC1: 192.168.1.1", "PC2: 192.168.1.2", "Subnet mask: 255.255.255.0", "Now they can communicate!"],
        "bg": "D5F5E3"
    },
    {
        "num": 23,
        "title": "Step-by-Step Tutorial: Testing with Ping",
        "script": "Let us test if our network works! Click on PC1. Go to the Desktop tab. Click Command Prompt. This opens a black window, like the command line. Type: ping 192.168.1.2 and press Enter. If you see 'Reply from 192.168.1.2' that means it works! The two computers can talk to each other. If you see 'Request timed out' something is wrong. Check your cables and IP addresses. Congratulations! You just built your first network!",
        "music": "Celebration music - 'Victory' style fanfare",
        "animation": "Success checkmark animation placeholder",
        "emotion": "Celebratory - big accomplishment!",
        "bullets": ["Click PC1 > Desktop > Command Prompt", "Type: ping 192.168.1.2", "Press Enter", "'Reply from...' = SUCCESS!", "'Request timed out' = check settings", "You built your first network!"],
        "bg": "FDEBD0"
    },
    {
        "num": 24,
        "title": "Common Problem 1: Packet Tracer is NOT the Real Internet",
        "script": "Now let us talk about common mistakes and things people get confused about. Number one: Packet Tracer is NOT the real internet! It is just a simulator. A pretend network. If you build a network in Packet Tracer, it does not actually connect to the real internet. You cannot open YouTube inside Packet Tracer. You cannot browse Facebook. It is like playing with toy cars. They look like real cars, but you cannot drive them to the mall!",
        "music": "Attention-getting music - 'Important Notice' style",
        "animation": "Warning sign with X mark animation",
        "emotion": "Clear and corrective - busting myths",
        "bullets": ["NOT the real internet!", "Just a simulator / pretend network", "Cannot browse real websites inside it", "Cannot open YouTube or Facebook", "Like toy cars - look real but cannot drive", "It is for LEARNING, not for actual internet use"],
        "bg": "FFFFFF"
    },
    {
        "num": 25,
        "title": "Common Problem 2: Cannot Install Apps Inside",
        "script": "Number two: You cannot install real applications inside Packet Tracer. Some students try to install antivirus or web browsers on the virtual computers. That does not work! The computers inside Packet Tracer are not real computers. They are just icons that simulate network behavior. You can only use the tools that Packet Tracer gives you, like the Command Prompt and the web browser simulation. But these are very limited compared to real applications.",
        "music": "Clarification music - 'Clear Thinking' style",
        "animation": "Crossed-out app icons animation",
        "emotion": "Patient correction - common mistake to fix",
        "bullets": ["Cannot install real software inside!", "Virtual computers are just icons", "They simulate network behavior only", "No antivirus, no Chrome, no apps", "Only Packet Tracer built-in tools work", "Command Prompt and basic web sim only"],
        "bg": "D6EAF8"
    },
    {
        "num": 26,
        "title": "Common Problem 3: Not All Commands Work",
        "script": "Number three: Not all Cisco commands work in Packet Tracer. Real Cisco routers have thousands of commands. Packet Tracer only supports the most common ones. If you type a command and it says 'Invalid input' it might not be a real error. The command might just not be available in Packet Tracer. This is normal! For your studies and exams, the important commands are all there. But if you want to do very advanced things, you might need GNS3 or real equipment.",
        "music": "Understanding music - 'Now I See' style",
        "animation": "Command list with some highlighted, some grayed out",
        "emotion": "Reassuring - this is normal, do not panic",
        "bullets": ["Not ALL Cisco commands are available", "'Invalid input' might mean: not supported", "This is normal! Do not panic!", "Important exam commands all work", "For advanced stuff: use GNS3 or real routers", "Packet Tracer covers the basics well"],
        "bg": "D5F5E3"
    },
    {
        "num": 27,
        "title": "Common Problem 4: PT is NOT the Same as GNS3 or EVE-NG",
        "script": "Number four: Packet Tracer is NOT the same as GNS3 or EVE-NG. What are those? GNS3 and EVE-NG are also network simulators, but they are more advanced. They run actual router operating systems. They need a powerful computer with lots of RAM. Packet Tracer is lighter and easier. It is perfect for beginners and CCNA students. If someone tells you to use GNS3 instead, know that they are different tools for different levels. Start with Packet Tracer!",
        "music": "Comparison music - 'Side by Side' style",
        "animation": "Three tools shown side by side with comparison",
        "emotion": "Informative and guiding - helping choose the right tool",
        "bullets": ["Packet Tracer is for beginners", "GNS3 = more advanced, needs powerful PC", "EVE-NG = even more advanced", "PT is lighter and easier to use", "Perfect for CCNA students", "Start with PT, move up later if needed"],
        "bg": "FDEBD0"
    },
    {
        "num": 28,
        "title": "Common Problem 5: You Need a NetAcad Account",
        "script": "Number five: You need a Cisco NetAcad account to download Packet Tracer. You cannot just find it on any random website. Do not download it from unofficial sites! Those might have viruses. The official way is to go to netacad.com, create a free account, and download from there. Some people think you need to pay. No! It is free. But you must have an account. Remember: netacad.com is the only safe place to get it.",
        "music": "Safety advisory music - 'Be Careful' style",
        "animation": "Lock and key / security icon animation",
        "emotion": "Protective and advisory - keeping students safe",
        "bullets": ["MUST have Cisco NetAcad account", "Go to netacad.com only!", "Do NOT download from random sites", "Random sites might have viruses!", "It is FREE - no payment needed", "Just create an account and download"],
        "bg": "FFFFFF"
    },
    {
        "num": 29,
        "title": "Common Problem 6: Always SAVE Your Work!",
        "script": "Number six: Always save your work! This is so important. Many students spend one hour building a beautiful network. Then their computer crashes. Or they close Packet Tracer without saving. And everything is gone! Always press Ctrl+S often. Save your file in a folder you can find. Name your files clearly, like 'Lab1-MyNetwork.pkt'. Do not name it 'untitled' or 'asdfgh'. You will thank yourself later when you need to find it again!",
        "music": "Urgent reminder music - 'Do Not Forget' style",
        "animation": "Flashing save icon with file disappearing dramatization",
        "emotion": "Urgent and caring - preventing heartbreak",
        "bullets": ["ALWAYS press Ctrl+S often!", "Students lose hours of work!", "Computer crashes happen!", "Save in a folder you can find", "Name files clearly: 'Lab1-MyNetwork.pkt'", "Do NOT name it 'untitled' or random letters"],
        "bg": "D6EAF8"
    },
    {
        "num": 30,
        "title": "Common Problem 7: Cable Types Matter!",
        "script": "Number seven: Cable types matter! You cannot use any cable for any connection. There are rules. Copper straight-through cable: use this to connect DIFFERENT devices. Like a PC to a switch, or a switch to a router. Copper cross-over cable: use this to connect SIMILAR devices. Like PC to PC, or switch to switch. If you use the wrong cable, the connection will not work! The dots will stay red. If you are confused, use the automatic cable tool. It picks the right one for you.",
        "music": "Attention to detail music - 'Precision' style",
        "animation": "Cable types shown with correct/incorrect examples",
        "emotion": "Detailed and precise - this matters!",
        "bullets": ["Straight-through: DIFFERENT devices (PC to Switch)", "Cross-over: SIMILAR devices (PC to PC)", "Wrong cable = red dots = no connection!", "Use automatic cable tool if confused", "It picks the right cable for you", "Fiber optic: for very fast, long connections"],
        "bg": "D5F5E3"
    },
    {
        "num": 31,
        "title": "Common Problem 8: Devices Need IP Addresses",
        "script": "Number eight: Devices need IP addresses. They do not get them for free! When you place a new computer in Packet Tracer, it has no IP address. It cannot talk to anyone. You must either set the IP address manually, or set up a DHCP server to give addresses automatically. Many students connect everything with cables but forget the IP addresses. Then they wonder why ping does not work. Remember: cables are like roads, but IP addresses are like house numbers. You need both!",
        "music": "Lightbulb moment music - 'Aha!' style",
        "animation": "Houses without numbers analogy animation",
        "emotion": "Enlightening - making the concept click",
        "bullets": ["New devices have NO IP address!", "You must set it manually OR use DHCP", "Without IP = cannot communicate", "Cables = roads, IP addresses = house numbers", "Need BOTH to send data", "Most common reason ping fails!"],
        "bg": "FDEBD0"
    },
    {
        "num": 32,
        "title": "Common Problem 9: Simulation vs Realtime Mode Confusion",
        "script": "Number nine: Simulation mode vs Realtime mode confusion. Some students accidentally switch to Simulation mode and then think their network is broken because nothing seems to happen. In Simulation mode, you must click the Play button or Capture/Forward button to make packets move. If you do not click, nothing happens! If your network seems frozen, check if you are in Simulation mode. Switch back to Realtime mode for normal operation. Both modes are useful, but know which one you are in!",
        "music": "Clarity music - 'Now I Understand' style",
        "animation": "Mode toggle animation showing difference",
        "emotion": "Clarifying - solving a common frustration",
        "bullets": ["Simulation mode: must click Play!", "Nothing moves by itself in Simulation", "Network seems frozen? Check your mode!", "Realtime = automatic, normal speed", "Simulation = manual, step by step", "Know which mode you are in!"],
        "bg": "FFFFFF"
    },
    {
        "num": 33,
        "title": "Common Problem 10: Old Files May Not Open in New Versions",
        "script": "Number ten: Old Packet Tracer files may not open in new versions! Packet Tracer gets updated every year or so. A file made in version 7 might not open perfectly in version 8. And a file from version 8 will definitely NOT open in version 7. Always check which version your school uses. If your teacher gives you a file, make sure your Packet Tracer version matches. If you get an error opening a file, version mismatch is probably the reason. Keep your Packet Tracer updated!",
        "music": "Warning advisory music - 'Heads Up' style",
        "animation": "Version numbers with compatibility arrows",
        "emotion": "Practical warning - saving frustration ahead of time",
        "bullets": ["Old .pkt files may not open in new versions", "New files will NOT open in old versions!", "Check which version your school uses", "Match your version to your teacher's", "Error opening file? Check version!", "Keep Packet Tracer updated"],
        "bg": "D6EAF8"
    },
    {
        "num": 34,
        "title": "Tips and Tricks: Work Smarter",
        "script": "Let me share some tips to work smarter in Packet Tracer. Tip one: Use the Note tool to label parts of your network. It helps you remember what each section does. Tip two: Use Place Note to add descriptions. Tip three: Ctrl+Z is undo. If you make a mistake, undo it! Tip four: Use copy and paste to duplicate devices quickly. Tip five: Organize your workspace. Put routers at the top, switches in the middle, and PCs at the bottom. It looks cleaner and is easier to troubleshoot.",
        "music": "Tips and hacks music - 'Life Hack' style",
        "animation": "Quick tip cards flipping in",
        "emotion": "Sharing secrets - insider knowledge",
        "bullets": ["Use Note tool to label network parts", "Ctrl+Z = undo mistakes", "Copy and paste to duplicate devices", "Organize: routers top, switches middle, PCs bottom", "Save often (Ctrl+S)", "Use descriptive device names"],
        "bg": "D5F5E3"
    },
    {
        "num": 35,
        "title": "Tips and Tricks: Study Effectively",
        "script": "More tips for studying with Packet Tracer. Tip one: Follow along with your lessons. If your teacher explains subnetting, build it in Packet Tracer. Tip two: Break complex networks into small pieces. Do not try to build everything at once. Tip three: Use Simulation mode to understand problems. If ping fails, watch where the packet gets stuck. Tip four: Practice the labs from Cisco NetAcad. They give you step-by-step labs to complete. Tip five: Teach someone else! The best way to learn is to explain it to a friend.",
        "music": "Study motivation music - 'Keep Going' style",
        "animation": "Study tips checklist with checkmarks appearing",
        "emotion": "Motivating and supportive - you can do this!",
        "bullets": ["Follow along with your teacher's lessons", "Break complex networks into small pieces", "Use Simulation mode to find problems", "Practice NetAcad labs", "Teach a friend - best way to learn!", "Do not give up - networking takes practice"],
        "bg": "FDEBD0"
    },
    {
        "num": 36,
        "title": "Bonus: Packet Tracer for CCNA Exam Preparation",
        "script": "Here is a bonus tip! If you are preparing for the CCNA exam, Packet Tracer is your best friend. The CCNA exam tests your ability to configure networks. Many exam questions are based on things you can practice in Packet Tracer. Practice configuring routers, setting up VLANs, configuring DHCP, and troubleshooting connectivity. The more you practice in Packet Tracer, the more confident you will be in the exam. Many Filipino IT professionals passed CCNA by practicing in Packet Tracer!",
        "music": "Achievement and goals music - 'Reach Higher' style",
        "animation": "Certificate/achievement badge animation",
        "emotion": "Aspirational - connecting to career goals",
        "bullets": ["Packet Tracer = CCNA best friend", "CCNA tests network configuration skills", "Practice: routers, VLANs, DHCP", "Practice troubleshooting connectivity", "More practice = more confidence", "Many Filipinos passed CCNA with PT!"],
        "bg": "FFFFFF"
    },
    {
        "num": 37,
        "title": "Bonus: Career Paths with Networking Skills",
        "script": "Learning Packet Tracer can open doors to great careers! In the Philippines, IT professionals with networking skills are in high demand. You could become a Network Administrator, managing networks for companies. Or a Network Engineer, designing new networks. Or a Systems Administrator. Or even a Cybersecurity Analyst. These jobs pay well and are available in many cities. BPO companies, banks, hospitals, and government offices all need networking professionals. Your Packet Tracer skills are the first step!",
        "music": "Inspirational career music - 'Bright Future' style",
        "animation": "Career path infographic animation",
        "emotion": "Inspiring and forward-looking - showing the future",
        "bullets": ["Network Administrator", "Network Engineer", "Systems Administrator", "Cybersecurity Analyst", "High demand in the Philippines!", "BPOs, banks, hospitals need you"],
        "bg": "D6EAF8"
    },
    {
        "num": 38,
        "title": "Summary: What We Learned Today",
        "script": "Let us review what we learned today! We learned that Cisco Packet Tracer is a free network simulator. We learned how to download it from netacad.com. We saw the workspace and device categories. We looked at six Philippine scenarios where networking helps communities. We did a step-by-step tutorial building our first network. We talked about ten common mistakes to avoid. And we shared tips and tricks for studying. That is a lot of learning! You should be proud of yourself.",
        "music": "Recap and reflection music - 'Looking Back' style",
        "animation": "Summary bullet points with checkmarks",
        "emotion": "Reflective and proud - celebrating the journey",
        "bullets": ["Packet Tracer = free network simulator", "Download from netacad.com", "Philippine scenarios show real-world use", "Built our first network together!", "10 common mistakes to avoid", "Tips and tricks for success"],
        "bg": "D5F5E3"
    },
    {
        "num": 39,
        "title": "Questions and Answers",
        "script": "Now it is time for your questions! Do you have anything you want to ask? Maybe you are confused about something we discussed. Maybe you want to know more about a specific topic. Maybe you want me to show you something in Packet Tracer. There are no silly questions! Every question helps you learn. And if you think of questions later, you can always reach out to your instructor or post in the Cisco NetAcad forums. The community is very helpful!",
        "music": "Open and welcoming music - 'Friendly Chat' style",
        "animation": "Question mark animations floating up",
        "emotion": "Open and welcoming - every question is valid",
        "bullets": ["Ask anything!", "No silly questions exist", "Want to see something demonstrated?", "Confused about a topic? Ask!", "Cisco NetAcad forums are helpful too", "Community is always ready to help"],
        "bg": "FDEBD0"
    },
    {
        "num": 40,
        "title": "Thank You and Happy Networking!",
        "script": "Thank you so much for joining us today! I hope you learned a lot about Cisco Packet Tracer. Remember, the best way to learn networking is to practice. Open Packet Tracer, build networks, break things, fix things, and have fun! Networking is like a puzzle, and every time you solve it, you get better. Keep learning, keep practicing, and one day you will build real networks that help real people. Maraming salamat and happy networking, everyone!",
        "music": "Celebration and farewell - 'Happy Ending' by Bensound",
        "animation": "Confetti and thank you message with fade out",
        "emotion": "Grateful and celebratory - warm farewell",
        "bullets": ["Thank you for joining!", "Practice is the best teacher", "Open PT and build networks", "Break things, fix things, have fun!", "Keep learning every day", "Maraming salamat! Happy networking!"],
        "bg": "FFFFFF"
    },
]


def generate_word_document():
    """Generate the Word document with the webinar script."""
    doc = Document()

    # Title
    title = doc.add_heading("Cisco Packet Tracer Webinar Script", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(
        "Complete Narration Script for 1-Hour Webinar\n"
        "Language: Simple English (explain like talking to a 5-year-old)\n"
        "Total Slides: 40 | Estimated Time: ~2 minutes per slide"
    )
    doc.add_paragraph("")  # spacer

    for slide in SLIDES:
        # Slide heading
        doc.add_heading(f"Slide {slide['num']}: {slide['title']}", level=1)

        # Script/Narration
        doc.add_heading("Script (What to Say):", level=2)
        p = doc.add_paragraph(slide["script"])
        p.style.font.size = Pt(12)

        # Music suggestion
        doc.add_heading("Background Music:", level=2)
        doc.add_paragraph(slide["music"])

        # Animation suggestion
        doc.add_heading("Animation/Transition:", level=2)
        doc.add_paragraph(slide["animation"])

        # Emotion/delivery
        doc.add_heading("Emotion/Delivery Style:", level=2)
        doc.add_paragraph(slide["emotion"])

        # Bullet points for reference
        doc.add_heading("Key Points on Slide:", level=2)
        for bullet in slide["bullets"]:
            doc.add_paragraph(bullet, style="List Bullet")

        # Separator
        doc.add_paragraph("_" * 60)
        doc.add_paragraph("")  # spacer

    output_path = "webinar/Cisco_Packet_Tracer_Webinar_Script.docx"
    doc.save(output_path)
    print(f"Word document saved: {output_path}")


def generate_powerpoint():
    """Generate the PowerPoint presentation."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Color map for backgrounds
    bg_colors = {
        "D6EAF8": PptRGB(0xD6, 0xEA, 0xF8),  # light blue
        "D5F5E3": PptRGB(0xD5, 0xF5, 0xE3),  # light green
        "FDEBD0": PptRGB(0xFD, 0xEB, 0xD0),  # light orange
        "FFFFFF": PptRGB(0xFF, 0xFF, 0xFF),  # white
    }

    for slide_data in SLIDES:
        # Use blank layout
        slide_layout = prs.slide_layouts[6]  # Blank
        slide = prs.slides.add_slide(slide_layout)

        # Set background color
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = bg_colors.get(slide_data["bg"], PptRGB(0xFF, 0xFF, 0xFF))

        # Add title text box
        left = Inches(0.5)
        top = Inches(0.3)
        width = Inches(12.3)
        height = Inches(1.2)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"Slide {slide_data['num']}: {slide_data['title']}"
        p.font.size = PptPt(36)
        p.font.bold = True
        p.font.color.rgb = PptRGB(0x1A, 0x23, 0x7E)
        p.alignment = PP_ALIGN.LEFT

        # Add bullet points
        left = Inches(0.8)
        top = Inches(1.8)
        width = Inches(11.5)
        height = Inches(5.0)
        txBox2 = slide.shapes.add_textbox(left, top, width, height)
        tf2 = txBox2.text_frame
        tf2.word_wrap = True

        for i, bullet in enumerate(slide_data["bullets"]):
            if i == 0:
                p2 = tf2.paragraphs[0]
            else:
                p2 = tf2.add_paragraph()
            p2.text = bullet
            p2.font.size = PptPt(24)
            p2.font.color.rgb = PptRGB(0x2C, 0x3E, 0x50)
            p2.space_after = PptPt(12)
            p2.level = 0

        # Add speaker notes
        notes_slide = slide.notes_slide
        notes_tf = notes_slide.notes_text_frame
        notes_tf.text = (
            f"SCRIPT: {slide_data['script']}\n\n"
            f"MUSIC: {slide_data['music']}\n"
            f"ANIMATION: {slide_data['animation']}\n"
            f"EMOTION: {slide_data['emotion']}"
        )

    output_path = "webinar/Cisco_Packet_Tracer_Presentation.pptx"
    prs.save(output_path)
    print(f"PowerPoint saved: {output_path}")


if __name__ == "__main__":
    print("Generating Cisco Packet Tracer Webinar materials...")
    print(f"Total slides: {len(SLIDES)}")
    generate_word_document()
    generate_powerpoint()
    print("Done! Both files are in the webinar/ folder.")
