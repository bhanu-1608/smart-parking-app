import re

with open(r'C:\smart_parking_app\generate_internship_report.py', 'r', encoding='utf-8') as f:
    code = f.read()

# Update TOC entries
toc_pattern = r'toc_entries = \[.*?\]'
new_toc = '''toc_entries = [
        ("Certificate from College", "ii"),
        ("Internship Completion Certificate", "iii"),
        ("Acknowledgement", "iv"),
        ("Abstract", "v"),
        ("List of Figures", "viii"),
        ("List of Tables", "ix"),
        ("1. INTRODUCTION", "1"),
        ("   1.1 About the Internship", "1"),
        ("   1.2 About the Host Organization", "1"),
        ("   1.3 Project Overview: DeccanPulse", "1"),
        ("   1.4 Problem Statement", "2"),
        ("   1.5 Motivation and Objectives", "2"),
        ("   1.6 Scope and Delimitations", "2"),
        ("   1.7 Organization of the Report", "3"),
        ("2. LITERATURE REVIEW AND THEORETICAL FOUNDATION", "4"),
        ("   2.1 Smart Parking Paradigms and Urban Mobility Systems", "4"),
        ("   2.2 Real-Time Communication: Polling versus WebSockets and SSE", "4"),
        ("   2.3 Geospatial Mathematics: The Haversine Formula and Coordinate Projection", "4"),
        ("   2.4 Multi-Deck Optimization and Unbundled EV Tariff Dynamics", "5"),
        ("   2.5 Web Cartography: Custom Lightweight SVG versus Third-Party Map SDKs", "5"),
        ("   2.6 Full-Stack Python Frameworks: ASGI Event-Loops with FastAPI", "5"),
        ("   2.7 Comparative Analysis: Conventional Parking versus DeccanPulse Platform", "6"),
        ("3. INTERNSHIP OBJECTIVES AND OPERATIONAL ROADMAP", "7"),
        ("   3.1 Academic and Professional Competency Objectives", "7"),
        ("   3.2 System Engineering Goals and Quality Attributes", "7"),
        ("   3.3 Learning and Technical Milestones", "7"),
        ("   3.4 Six-Week Chronological Timeline and Sprint Breakdown", "7"),
        ("   3.5 Expected Engineering Outcomes and Deliverables", "8"),
        ("4. TECHNOLOGY STACK, ARCHITECTURE TOOLING AND REQUIREMENTS", "9"),
        ("   4.1 Overview of the Decoupled System Architecture", "9"),
        ("   4.2 Backend Technology Stack and Dependency Specifications", "9"),
        ("   4.3 Frontend Client Architecture and Reactive Component Tree", "9"),
        ("   4.4 Data Persistence, Schemas and In-Memory Concurrency Store", "10"),
        ("   4.5 Engineering Development and Hardware Runtime Requirements", "10"),
        ("   4.6 Core Engineering Competencies Acquired", "11"),
        ("5. ENGINEERING CONTRIBUTIONS AND INNOVATIONS", "12"),
        ("   5.1 Summary of Key Technical Achievements", "12"),
        ("   5.2 Granular Technical Contributions", "12"),
        ("   5.3 Innovation 1: Multi-Deck Bay Allocation and Unbundled EV Charging Engine", "12"),
        ("   5.4 Innovation 2: True North Vector Compass and Floating Cultural Legend", "13"),
        ("   5.5 Innovation 3: Turn-by-Turn Routing and GPS Drive Telemetry Simulator", "13"),
        ("   5.6 Innovation 4: Dual Heritage Design System (Charminar Gold & Nizami Onyx)", "13"),
        ("   5.7 Engineering Challenges Encountered and Technical Resolutions", "14"),
        ("   5.8 Master Deliverables Catalogue", "14"),
        ("6. SYSTEM DESIGN, MODELING AND IMPLEMENTATION", "15"),
        ("   6.1 Ten-Step End-to-End Operational Lifecycle", "15"),
        ("   6.2 High-Level Architecture and Data Flow Pipeline", "15"),
        ("   6.3 Relational Database Schema and SQLAlchemy Entity Models", "16"),
        ("   6.4 Role-Based Access Control and Permission Security Matrix", "16"),
        ("   6.5 Real-Time IoT Telemetry Simulation and Background Task Daemon", "17"),
        ("   6.6 Geospatial Proximity Matchmaking and Dynamic Path Routing", "17"),
        ("   6.7 Real-Time Event Protocol and Telemetry Event Catalogue", "18"),
        ("   6.8 REST API Endpoint Directory and Interface Specification", "18"),
        ("   6.9 Frontend Single Page Application State and Navigation Routes", "19"),
        ("   6.10 Gate Pass Lifecycle and Dynamic Optical QR Verification", "19"),
        ("   6.11 Visual Artifacts and Graphical User Interface Verification", "20"),
        ("7. TESTING, RESULTS AND PERFORMANCE ANALYSIS", "26"),
        ("   7.1 Comprehensive Testing Setup and Verification Methodology", "26"),
        ("   7.2 Functional and Integration Test Cases (TC-01 through TC-14)", "26"),
        ("   7.3 Multi-Category Pricing and Dynamic Tariff Mathematical Verification", "27"),
        ("   7.4 Real-Time Telemetry Latency Benchmarking", "28"),
        ("   7.5 Codebase Metrics, Complexity and Structural Distribution", "28"),
        ("   7.6 Code Review Findings and Security Hardening Recommendations", "29"),
        ("   7.7 Key Empirical Observations and Operational Findings", "29"),
        ("8. SOURCE CODE ARCHITECTURE AND KEY LISTINGS", "30"),
        ("   8.1 Listing 8.1: FastAPI Asynchronous Lifespan and IoT Heartbeat Loop", "30"),
        ("   8.2 Listing 8.2: Dynamic Unbundled Pricing Calculation Engine", "30"),
        ("   8.3 Listing 8.3: Geospatial Haversine & Screen Coordinate Transformation", "31"),
        ("   8.4 Listing 8.4: Real-Time GPS Vehicle Navigation and Step Simulator", "31"),
        ("   8.5 Listing 8.5: Reactive SVG Map Rendering and Theme Controller", "32"),
        ("   8.6 Listing 8.6: Gate Pass Cryptographic QR Generation and Verification", "32"),
        ("9. INDUSTRIAL APPLICATIONS, SCALABILITY AND FUTURE SCOPE", "34"),
        ("   9.1 Real-World Metropolitan Deployment Scenarios", "34"),
        ("   9.2 High-Throughput Cloud Scaling and Containerization Roadmap", "34"),
        ("   9.3 Product Enhancements: Computer Vision ANPR and FASTag Integration", "35"),
        ("   9.4 Privacy, Ethics and Municipal Urban Governance", "35"),
        ("10. CONCLUSION AND RETROSPECTIVE SUMMARY", "36"),
        ("   10.1 Project Conclusion", "36"),
        ("   10.2 Retrospective Learning Outcomes", "36"),
        ("   10.3 Final Concluding Statement", "36"),
        ("REFERENCES", "37")
    ]'''
code = re.sub(toc_pattern, new_toc, code, flags=re.DOTALL)

# Update Figures data
fig_pattern = r'figures_data = \[.*?\]'
new_fig = '''figures_data = [
        ("Fig. 1.1", "Ten-step operational lifecycle of a parking session in DeccanPulse", "3"),
        ("Fig. 3.1", "Six-week structured internship roadmap and sprint execution chart", "8"),
        ("Fig. 6.1", "High-level decoupled system architecture of the DeccanPulse platform", "15"),
        ("Fig. 6.2", "Entity-Relationship diagram of the five core database models", "16"),
        ("Fig. 6.3", "Sequence diagram of gate pass reservation, QR verification and bay checkout", "19"),
        ("Fig. 6.4", "Interactive GIS Map Canvas in Light Mode with True North Compass and Legend", "20"),
        ("Fig. 6.5", "Intelligent Route Guidance and Step-by-Step Turn Navigation Pane", "20"),
        ("Fig. 6.6", "Real-Time GPS Drive Simulation with Live Vehicle Telemetry and Path Tracker", "21"),
        ("Fig. 6.7", "Hyderabad Parking Hubs Directory and Multi-Zone Occupancy Matrix", "21"),
        ("Fig. 6.8", "User Vehicle Garage Management with Instant Vehicle Fast-Switch Selector", "22"),
        ("Fig. 6.9", "Multi-Deck Bay Allocation Modal with Duration-Tiered Pricing Breakdown", "23"),
        ("Fig. 6.10", "Digital Boarding Pass Gate Pass with Real-Time Dynamic QR Code", "24"),
        ("Fig. 6.11", "Deccan Heritage Dark Mode utilizing Hyderabad Charminar & Onyx Cultural Palettes", "25"),
        ("Fig. 7.1", "Real-time telemetry heartbeat sensor oscillation across 10 hubs", "28"),
        ("Fig. 7.2", "Distribution of source code across architectural modules", "28")
    ]'''
code = re.sub(fig_pattern, new_fig, code, flags=re.DOTALL)

# Update Tables data
tbl_pattern = r'tables_data = \[.*?\]'
new_tbl = '''tables_data = [
        ("Table 2.1", "Conventional On-Street Parking Hunting versus DeccanPulse Smart Orchestration", "6"),
        ("Table 3.1", "Six-Week Structured Internship Roadmap and Milestone Log (1 May – 15 June 2026)", "8"),
        ("Table 4.1", "Backend Software Dependencies and Python 3.12 Runtime Specifications", "9"),
        ("Table 4.2", "Frontend Client-Side Technologies and Component Libraries", "10"),
        ("Table 4.3", "Engineering Development, Runtime Environment and Hardware Specifications", "11"),
        ("Table 5.1", "Engineering Challenges Encountered and Technical Resolutions", "14"),
        ("Table 5.2", "Master Catalogue of Internship Production Deliverables", "14"),
        ("Table 6.1", "Relational Database Schema and SQLAlchemy Entity Definitions", "16"),
        ("Table 6.2", "Role-Based Access Control and Permission Security Matrix", "17"),
        ("Table 6.3", "Real-Time Telemetry and WebSocket / Server-Sent Event (SSE) Protocol Catalog", "18"),
        ("Table 6.4", "Complete REST API Endpoint Directory (24 Production Endpoints)", "18"),
        ("Table 6.5", "Frontend Single Page Application State and View Navigation Directory", "19"),
        ("Table 6.6", "Vehicle Category Bay Surcharge Multipliers and Base Tariffs", "19"),
        ("Table 7.1", "Functional and Integration Test Cases and Execution Outcomes (TC-01 – TC-14)", "27"),
        ("Table 7.2", "Multi-Category Pricing Verification Across Vehicle Classes and Durations", "28"),
        ("Table 7.3", "Real-Time Sensor Telemetry Latency Benchmarks Under Simulated Network Loads", "28"),
        ("Table 7.4", "Source Codebase Metric Breakdown by Architectural Module", "29"),
        ("Table 7.5", "Code Review Findings, Vulnerability Assessments and Hardening Recommendations", "29")
    ]'''
code = re.sub(tbl_pattern, new_tbl, code, flags=re.DOTALL)

# Ensure List of Tables starts on its own page
if 'builder.add_page_break()\n    p_lot = doc.add_paragraph()' not in code:
    code = code.replace(
        'p_lot = doc.add_paragraph()\n    p_lot.alignment = WD_ALIGN_PARAGRAPH.CENTER',
        'builder.add_page_break()\n    p_lot = doc.add_paragraph()\n    p_lot.alignment = WD_ALIGN_PARAGRAPH.CENTER'
    )

with open(r'C:\smart_parking_app\generate_internship_report.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("Updated generate_internship_report.py successfully via update_toc.py!")
