import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# All college data collected from official websites
college_data = [
    {
        "name": "American University of Health Sciences",
        "website": "https://www.auhs.edu",
        "courses": [
            ("Bachelor of Science in Nursing (BSN)", "Bachelor's Degree"),
            ("RN to BSN Program", "Bachelor's Degree"),
            ("LVN 45-Unit Option (LVN to BSN)", "Bachelor's Degree"),
            ("Master of Science in Nursing (MSN)", "Master's Degree"),
            ("Bachelor of Science in Pharmaceutical Sciences (BSPS)", "Bachelor's Degree"),
            ("Doctor of Pharmacy (PharmD)", "Doctoral Degree"),
            ("Master of Science in Clinical Research (MSCR)", "Master's Degree"),
            ("Pharmacy Technician", "Extension Program"),
            ("Nursing Assistant Training Program", "Extension Program"),
        ]
    },
    {
        "name": "Artistic Nails & Beauty Academy",
        "website": "https://www.artistic.edu",
        "courses": [
            ("Cosmetology", "Diploma/Certificate"),
            ("Nail Program (Nail Technology)", "Diploma/Certificate"),
            ("Skincare Program", "Diploma/Certificate"),
            ("Full Specialist Program (Nails + Skincare)", "Diploma/Certificate"),
            ("Barbering", "Diploma/Certificate"),
            ("Massage Therapy", "Diploma/Certificate"),
            ("Spa Therapy", "Diploma/Certificate"),
        ]
    },
    {
        "name": "Avalon Institute",
        "website": "https://avalon.edu",
        "courses": [
            ("Cosmetology", "Diploma/Certificate"),
            ("Esthetics", "Diploma/Certificate"),
            ("Makeup Artistry", "Advanced Class"),
            ("Eyelash Extensions", "Advanced Class"),
            ("Hair Extension Training", "Advanced Class"),
            ("Massage Therapy", "Diploma/Certificate"),
            ("Barbering", "Diploma/Certificate"),
            ("Student Instructor (Cosmetology/Esthetics Instructor Training)", "Diploma/Certificate"),
        ]
    },
    {
        "name": "Charter College",
        "website": "https://chartercollege.edu",
        "courses": [
            ("Associate of Applied Science in Allied Health (Degree Completion)", "Associate's Degree - Healthcare"),
            ("Associate of Applied Science in Diagnostic Medical Sonography", "Associate's Degree - Healthcare"),
            ("Associate of Applied Science in Medical Billing and Coding", "Associate's Degree - Healthcare"),
            ("Associate of Applied Science in Nursing", "Associate's Degree - Healthcare"),
            ("Associate of Applied Science in Radiologic Technology", "Associate's Degree - Healthcare"),
            ("Associate of Applied Science in Respiratory Therapy", "Associate's Degree - Healthcare"),
            ("Advanced Certificate in Computed Tomography", "Advanced Certificate - Healthcare"),
            ("Advanced Certificate in Magnetic Resonance Imaging", "Advanced Certificate - Healthcare"),
            ("Advanced Certificate in Vascular Sonography", "Advanced Certificate - Healthcare"),
            ("Bachelor of Science in Health Care Administration (Degree Completion)", "Bachelor's Degree - Healthcare"),
            ("Bachelor of Science in Health Technology Management (Degree Completion)", "Bachelor's Degree - Healthcare"),
            ("Certificate in Dental Assisting", "Certificate - Healthcare"),
            ("Certificate in Health Unit Coordinator", "Certificate - Healthcare"),
            ("Certificate in Medical Assistant", "Certificate - Healthcare"),
            ("Certificate in Medical Office Administrative Assistant", "Certificate - Healthcare"),
            ("Certificate in Pharmacy Technician", "Certificate - Healthcare"),
            ("Diploma in Phlebotomy", "Diploma - Healthcare"),
            ("Associate of Applied Science in Business Administration", "Associate's Degree - Business"),
            ("Associate of Applied Science in Business Management", "Associate's Degree - Business"),
            ("Associate of Applied Science in Paralegal", "Associate's Degree - Business"),
            ("Bachelor of Science in Business Administration (Degree Completion)", "Bachelor's Degree - Business"),
            ("Associate of Applied Science in Cybersecurity", "Associate's Degree - Information Technology"),
            ("Bachelor of Science in Computer Information Systems", "Bachelor's Degree - Information Technology"),
            ("Associate of Applied Science in Applied Technology (Degree Completion)", "Associate's Degree - Trades"),
            ("Certificate in Heating, Ventilation, Air Conditioning, and Refrigeration (HVAC/R)", "Certificate - Trades"),
            ("Certificate in Welding", "Certificate - Trades"),
            ("Certificate in Veterinary Assistant", "Certificate - Veterinary"),
        ]
    },
    {
        "name": "CyberTex Institute of Technology",
        "website": "https://cybertex.edu",
        "courses": [
            ("Network Engineer Program (IT/Cybersecurity)", "Diploma/Certificate"),
            ("Medical Assistant Program", "Diploma/Certificate"),
            ("Vocational Nursing (LVN) Program", "Diploma/Certificate"),
            ("HVAC (Heating, Ventilation and Air Conditioning) Program", "Diploma/Certificate"),
        ]
    },
    {
        "name": "Educators of Beauty",
        "website": "https://www.educatorsofbeauty.com",
        "courses": [
            ("Cosmetology (1,500 hours)", "Diploma/Certificate"),
            ("Esthetics (750 hours)", "Diploma/Certificate"),
            ("Nail Technology (350 hours)", "Diploma/Certificate"),
            ("Cosmetology Teacher/Instructor Training (500 hours)", "Diploma/Certificate"),
            ("Cosmetology Continuing Education (CEU)", "Continuing Education"),
        ]
    },
    {
        "name": "Florida Christian University",
        "website": "https://www.floridachristianuniversity.edu",
        "courses": [
            ("Associate of Arts in Theology", "Associate's Degree"),
            ("Associate of Arts in Counseling", "Associate's Degree"),
            ("Bachelor of Arts in Theology", "Bachelor's Degree"),
            ("Bachelor of Arts in Counseling", "Bachelor's Degree"),
            ("Bachelor of Science in Education", "Bachelor's Degree"),
            ("Bachelor of Science in Business Administration", "Bachelor's Degree"),
            ("Master of Arts in Theology", "Master's Degree"),
            ("Master of Arts in Clinical Counseling", "Master's Degree"),
            ("Master of Arts in Marriage and Family Therapy", "Master's Degree"),
            ("Master of Arts in Coaching", "Master's Degree"),
            ("Master of Arts in Principled Education", "Master's Degree"),
            ("Master of Science in Education", "Master's Degree"),
            ("Master of Science in Business Administration", "Master's Degree"),
            ("Master of Science in Business Administration in Health Care Services", "Master's Degree"),
            ("Master of Science in Business Administration in Cognitive Neuroscience", "Master's Degree"),
            ("Doctor of Philosophy in Theology", "Doctoral Degree"),
            ("Doctor of Philosophy in Education", "Doctoral Degree"),
            ("Doctor of Philosophy in Clinical Counseling", "Doctoral Degree"),
            ("Doctor of Philosophy in Coaching", "Doctoral Degree"),
            ("Doctor of Philosophy in Business Administration", "Doctoral Degree"),
            ("Doctor of Philosophy in Business Administration in Financial Education", "Doctoral Degree"),
            ("Doctor of Philosophy in Business Administration in Health Care Services", "Doctoral Degree"),
            ("Post-Doctoral Diploma in Business Administration", "Post-Doctoral"),
            ("Post-Doctor in Education", "Post-Doctoral"),
            ("Certificate Courses (40 total including Coaching Fundamentals, Theology, Counseling, Education, Business)", "Certificate"),
        ]
    },
    {
        "name": "Gurnick Academy of Medical Arts",
        "website": "https://www.gurnick.edu",
        "courses": [
            ("A.O.S. in Respiratory Therapy", "Associate's Degree"),
            ("A.O.S. in Cardiac Ultrasound Technology", "Associate's Degree"),
            ("A.O.S. in Ultrasound Technology", "Associate's Degree"),
            ("A.O.S. in Vascular Ultrasound Technology", "Associate's Degree"),
            ("A.S. in Magnetic Resonance Imaging (MRI)", "Associate's Degree"),
            ("A.S. in Nuclear Medicine Technology", "Associate's Degree"),
            ("A.S. in Nursing (ADN)", "Associate's Degree"),
            ("A.S. in Nursing (LVN to RN)", "Associate's Degree"),
            ("A.S. in Occupational Therapy Assistant", "Associate's Degree"),
            ("A.S. in Physical Therapist Assistant", "Associate's Degree"),
            ("A.S. in Radiologic Technology (Track A)", "Associate's Degree"),
            ("A.S. in Radiologic Technology (Track B)", "Associate's Degree"),
            ("A.S. in Vocational Nursing", "Associate's Degree"),
            ("B.S. in Diagnostic Medical Imaging / Radiologic Sciences", "Bachelor's Degree"),
            ("B.S. in Nursing (BSN / LVN to BSN)", "Bachelor's Degree"),
            ("B.S. in Nursing (RN to BSN)", "Bachelor's Degree"),
            ("B.S. in Radiation Therapy", "Bachelor's Degree"),
            ("M.S. in Nursing (BSN to MSN)", "Master's Degree"),
            ("International Nurse Graduate Courses (Medical-Surgical, Maternal/Newborn, Pediatric, Mental Health Nursing)", "Continuing Education"),
            ("IV Therapy / Blood Withdrawal", "Continuing Education"),
        ]
    },
    {
        "name": "Health Staff Training Institute",
        "website": "https://hsti.com",
        "courses": [
            ("Clinical & Administrative Medical Assistant", "Diploma/Certificate"),
            ("Clinical Medical Assistant", "Diploma/Certificate"),
            ("Administrative Medical Assistant", "Diploma/Certificate"),
            ("Medical Front Office Assistant", "Diploma/Certificate"),
            ("Medical Billing & Coding", "Diploma/Certificate"),
            ("Advanced Medical Coding", "Diploma/Certificate"),
            ("Electronic Health Records", "Diploma/Certificate"),
            ("Phlebotomy Technician", "Diploma/Certificate"),
            ("Pharmacy Technician", "Diploma/Certificate"),
            ("Drug & Alcohol Counseling (SUD – Substance Use Disorder)", "Diploma/Certificate"),
            ("Computerized Office & Accounting", "Diploma/Certificate"),
            ("Introduction to Computer Technology", "Diploma/Certificate"),
        ]
    },
    {
        "name": "Healthcare Career College",
        "website": "https://healthcarecareercollege.edu",
        "courses": [
            ("Diagnostic Medical Sonography (Ultrasound) – Diploma", "Diploma"),
            ("Diagnostic Medical Sonography (Ultrasound) – Associate's Degree", "Associate's Degree"),
            ("Phlebotomy Technician", "Diploma/Certificate"),
            ("Dental Assistant", "Diploma/Certificate"),
            ("Physical Therapy Aide", "Diploma/Certificate"),
            ("Medical Assistant", "Diploma/Certificate"),
            ("Massage Therapy", "Diploma/Certificate"),
            ("Healthcare Management (Associate Degree)", "Associate's Degree"),
            ("Computerized Medical Biller and Coder", "Diploma/Certificate"),
            ("Nurse Assistant / Home Health Aide", "Diploma/Certificate"),
            ("Patient Care Technician", "Diploma/Certificate"),
        ]
    },
    {
        "name": "Heavy Equipment College of America",
        "website": "https://heavyequipmentcollege.edu",
        "courses": [
            ("Certificate of Heavy Equipment Operations – Level I", "Certificate"),
            ("Certificate of Heavy Equipment Operations – Level II", "Certificate"),
            ("Mobile Crane Operation", "Certificate"),
            ("Fixed Cab (Tower) Crane Operation", "Certificate"),
            ("Lattice Boom Crawler Crane Operation", "Certificate"),
            ("Heating and Air Technology (HVAC)", "Certificate"),
            ("Refrigeration Certification", "Certificate"),
            ("Horizontal Directional Drilling (HDD)", "Certificate"),
            ("Hybrid Degree Program – Heavy Equipment Operations with Estimating & Project Management", "Associate's Degree"),
        ]
    },
    {
        "name": "IBMC Intellitec",
        "website": "https://ibmc.edu / https://intellitec.edu",
        "courses": [
            # IBMC College
            ("Cosmetology (IBMC College)", "Diploma/Certificate"),
            ("Barbering (IBMC College)", "Diploma/Certificate"),
            ("Esthetics / Esthetician (IBMC College)", "Diploma/Certificate"),
            ("Therapeutic Massage (IBMC College)", "Diploma/Certificate"),
            ("Paralegal (IBMC College)", "Diploma/Certificate"),
            ("Clinical Medical Assisting – Diploma (IBMC College)", "Diploma"),
            ("Clinical Medical Assisting – Associate Degree (IBMC College)", "Associate's Degree"),
            ("Dental Assisting (IBMC College)", "Diploma/Certificate"),
            ("Pharmacy Technician (IBMC College)", "Diploma/Certificate"),
            ("Computer Systems Technician (IBMC College)", "Diploma/Certificate"),
            # IntelliTec College
            ("Automotive Technology (IntelliTec College)", "Diploma/Certificate"),
            ("Electrical Technology (IntelliTec College)", "Diploma/Certificate"),
            ("Refrigeration & HVAC (IntelliTec College)", "Diploma/Certificate"),
            ("Information Technology / Computer Systems Networking (IntelliTec College)", "Diploma/Certificate"),
            ("Medical Assistant (IntelliTec College)", "Diploma/Certificate"),
            ("Dental Assistant (IntelliTec College)", "Diploma/Certificate"),
            ("Medical Billing and Coding (IntelliTec College)", "Diploma/Certificate"),
            ("Cosmetology (IntelliTec College)", "Diploma/Certificate"),
            ("Massage Therapy (IntelliTec College)", "Diploma/Certificate"),
        ]
    },
    {
        "name": "ICPR Junior College",
        "website": "https://www.icprjc.edu",
        "courses": [
            ("B.S. in Medical Sonography with Specialty in Cardiovascular and Musculoskeletal", "Bachelor's Degree"),
            ("A.S. in Medical Sonography", "Associate's Degree"),
            ("A.S. in Medical Sonography (Online)", "Associate's Degree"),
            ("A.S. in Nursing", "Associate's Degree"),
            ("A.S. in Criminal Investigation and Forensic Sciences", "Associate's Degree"),
            ("A.S. in Diagnostic Radiological Technology", "Associate's Degree"),
            ("Associate in Business Administration – Health Services Coordination and Medical Billing", "Associate's Degree"),
            ("Gastronomy / Culinary Arts and Related Services", "Associate's Degree"),
            ("Practical Nursing", "Professional Certificate"),
            ("Pharmacy Technician", "Professional Certificate"),
            ("Medical Billing", "Professional Certificate"),
            ("Culinary Arts", "Professional Certificate"),
            ("Commercial Baking and Pastry Arts", "Professional Certificate"),
            ("Massage Therapist", "Professional Certificate"),
            ("Dental Technology / Dental Laboratory Technology", "Professional Certificate"),
            ("Digital Graphic Design and Web Page Development", "Professional Certificate"),
            ("Computer Repair / Computer Maintenance", "Professional Certificate"),
            ("Adult Patient Care", "Professional Certificate"),
            ("Early Education Assistant", "Professional Certificate"),
            ("Private Detective and Forensic Investigation", "Professional Certificate"),
            ("Homeland Security / Law Enforcement", "Professional Certificate"),
        ]
    },
    {
        "name": "Lincoln Technical Institute Inc",
        "website": "https://www.lincolntech.edu",
        "courses": [
            ("Automotive Technology / Automotive Service Technology", "Automotive"),
            ("Collision Repair and Refinishing", "Automotive"),
            ("Diesel Technology", "Automotive"),
            ("High Performance Automotive", "Automotive"),
            ("HVAC (Heating, Ventilation & Air Conditioning / Refrigeration)", "Skilled Trades"),
            ("Welding and Fabrication Technology", "Skilled Trades"),
            ("Electrical / Electronic Systems Technology", "Skilled Trades"),
            ("Electronic Engineering Technology", "Skilled Trades"),
            ("Advanced Manufacturing with Robotics (CNC Machining)", "Skilled Trades"),
            ("Drafting", "Skilled Trades"),
            ("Medical Assistant / Medical Assistant Technology", "Health Sciences"),
            ("Practical Nursing (LPN)", "Health Sciences"),
            ("Nursing Assistant / Patient Care", "Health Sciences"),
            ("Medical Coding and Billing / Medical Administration", "Health Sciences"),
            ("Medical Office Assistant", "Health Sciences"),
            ("Dental Assistant", "Health Sciences"),
            ("Pharmacy Technician", "Health Sciences"),
            ("Massage Therapy", "Health Sciences"),
            ("Computer and Network Support Technician", "Information Technology"),
            ("Networking / Network Communications and Information Systems", "Information Technology"),
            ("A+, Network+, Security+ Certification Preparation", "Information Technology"),
            ("Business Administration", "Business"),
            ("Criminal Justice", "Business"),
            ("Culinary Arts", "Business"),
            ("Cosmetology", "Beauty"),
        ]
    },
    {
        "name": "Lindenwood University",
        "website": "https://www.lindenwood.edu",
        "courses": [
            ("Accounting (BA, BS)", "Undergraduate"),
            ("Advertising and Public Relations (BA)", "Undergraduate"),
            ("Applied Science (BAS)", "Undergraduate"),
            ("Art and Design (BA)", "Undergraduate"),
            ("Art and Design with K-12 Education Certification (BA)", "Undergraduate"),
            ("Art History and Visual Culture (BA)", "Undergraduate"),
            ("Biology (BA, BS)", "Undergraduate"),
            ("Broadcast and Media Production (BA)", "Undergraduate"),
            ("Chemistry (BA, BS)", "Undergraduate"),
            ("Communication Studies (BA)", "Undergraduate"),
            ("Computer Science (BA, BS)", "Undergraduate"),
            ("Creative Writing (BA)", "Undergraduate"),
            ("Criminology and Criminal Justice (BA)", "Undergraduate"),
            ("Criminology and Criminal Justice with Law Enforcement Academy Emphasis (BA)", "Undergraduate"),
            ("Cybersecurity (BS)", "Undergraduate"),
            ("Dance (BA, BFA)", "Undergraduate"),
            ("Data Science (BS)", "Undergraduate"),
            ("Early Childhood Education (BA)", "Undergraduate"),
            ("Ecology and Evolutionary Biology (BS)", "Undergraduate"),
            ("Education (BA)", "Undergraduate"),
            ("Elementary Education (BA)", "Undergraduate"),
            ("English (BA)", "Undergraduate"),
            ("Environmental Science (BS)", "Undergraduate"),
            ("Exercise Science (BS)", "Undergraduate"),
            ("Fashion Business and Entrepreneurship (BS)", "Undergraduate"),
            ("Fashion Design and Technology (BFA)", "Undergraduate"),
            ("Finance (BA, BS)", "Undergraduate"),
            ("Game Design (BA)", "Undergraduate"),
            ("Health and Wellness (BS)", "Undergraduate"),
            ("History (BA)", "Undergraduate"),
            ("Human Services (BA)", "Undergraduate"),
            ("Information Technology (BS)", "Undergraduate"),
            ("International Business (BA, BS)", "Undergraduate"),
            ("Marketing (BA)", "Undergraduate"),
            ("Mathematics (BA, BS)", "Undergraduate"),
            ("Music (BA/BM)", "Undergraduate"),
            ("Music Performance (BA)", "Undergraduate"),
            ("Nursing RN to BSN Completion Track (BS)", "Undergraduate"),
            ("Photography (BA)", "Undergraduate"),
            ("Political Science (BA)", "Undergraduate"),
            ("Pre-Law (Pre-Professional track)", "Undergraduate"),
            ("Psychology (BS)", "Undergraduate"),
            ("Social Work (BSW)", "Undergraduate"),
            ("Sport Management (BA)", "Undergraduate"),
            ("Theatre (BA, BFA)", "Undergraduate"),
            ("Accountancy (MAcc)", "Graduate"),
            ("Advertising and Strategic Communications (MA)", "Graduate"),
            ("Art History and Visual Culture (MA)", "Graduate"),
            ("Art and Design (MA)", "Graduate"),
            ("Behavior Analysis (MA)", "Graduate"),
            ("Business Administration (MBA)", "Graduate"),
            ("Cybersecurity Management (MS)", "Graduate"),
            ("Education (MA, MAT, EdD)", "Graduate"),
            ("Fashion Business and Entrepreneurship (MS)", "Graduate"),
            ("Healthcare Administration (MHA)", "Graduate"),
            ("Higher Education (MA)", "Graduate"),
            ("Human Resource Management (MA)", "Graduate"),
            ("Information Technology Management (MS)", "Graduate"),
            ("Integrated School Library Media and Technology (MA)", "Graduate"),
            ("Interdisciplinary Media Arts (MFA)", "Graduate"),
            ("Leadership (MA, EdD)", "Graduate"),
            ("Music Education (MME)", "Graduate"),
            ("Nonprofit Administration (MA)", "Graduate"),
            ("Public Administration (MPA)", "Graduate"),
            ("Studio Art (MA)", "Graduate"),
            ("Writing (MFA)", "Graduate"),
            ("Graduate Certificate in Finance", "Graduate Certificate"),
            ("Graduate Certificate in Healthcare Administration", "Graduate Certificate"),
            ("Graduate Certificate in Human Resource Management", "Graduate Certificate"),
            ("Graduate Certificate in International Business", "Graduate Certificate"),
            ("Graduate Certificate in Leadership", "Graduate Certificate"),
            ("Graduate Certificate in Marketing", "Graduate Certificate"),
            ("Graduate Certificate in Project Management", "Graduate Certificate"),
            ("Graduate Certificate in Supply Chain Management", "Graduate Certificate"),
        ]
    },
    {
        "name": "Lu Ross Academy",
        "website": "https://www.lurossacademy.com",
        "courses": [
            ("Cosmetology (1,250-hour program)", "Diploma/Certificate"),
            ("Barbering", "Diploma/Certificate"),
            ("Barber Crossover", "Diploma/Certificate"),
            ("Aesthetics", "Diploma/Certificate"),
            ("Manicuring (600-hour program)", "Diploma/Certificate"),
            ("Beauty Makeup Artistry", "Diploma/Certificate"),
        ]
    },
    {
        "name": "New England College of Business",
        "website": "https://www.necb.edu",
        "courses": [
            ("Basic Accounting Certificate", "Undergraduate Certificate"),
            ("Intermediate Accounting Certificate", "Undergraduate Certificate"),
            ("Certificate in Digital Marketing", "Undergraduate Certificate"),
            ("Associate in Science in Business Administration", "Associate's Degree"),
            ("Bachelor of Science in Business Administration (BSBA)", "Bachelor's Degree"),
            ("Bachelor of Science in Digital Marketing (BSDM)", "Bachelor's Degree"),
            ("Bachelor of Science in International Business (BSIB)", "Bachelor's Degree"),
            ("Bachelor of Science in Quality Systems Management", "Bachelor's Degree"),
            ("Master of Business Administration (MBA)", "Master's Degree"),
            ("Master of Science in Finance (MSF)", "Master's Degree"),
            ("Master of Science in Business Ethics and Compliance (MBEC)", "Master's Degree"),
            ("Master of Healthcare Management (MHM)", "Master's Degree"),
            ("Master of Human Resource Management (MHRM)", "Master's Degree"),
            ("Master of Science in Quality Systems & Improvement Management (MSQSIM)", "Master's Degree"),
            ("Doctor of Business Administration (DBA) in Quality Systems & Improvement Management", "Doctoral Degree"),
            ("Graduate Certificate in Financial Management", "Graduate Certificate"),
            ("Graduate Certificate in Financial Planning", "Graduate Certificate"),
        ]
    },
    {
        "name": "Northwest Career College",
        "website": "https://www.northwestcareercollege.edu",
        "courses": [
            ("Business Administration (AAS)", "Associate's Degree"),
            ("Criminal Justice (AAS)", "Associate's Degree"),
            ("Healthcare Administration (AAS)", "Associate's Degree"),
            ("Paralegal Studies (AAS)", "Associate's Degree"),
            ("Radiography (AAS)", "Associate's Degree"),
            ("Business Administrative Assistant (Diploma)", "Diploma"),
            ("Criminal Justice Professional (Diploma)", "Diploma"),
            ("Dental Assistant (Diploma)", "Diploma"),
            ("Pharmacy Technician (Diploma)", "Diploma"),
            ("Dental Administrative Assistant (Certificate)", "Certificate"),
            ("Phlebotomy Technician (Certificate)", "Certificate"),
        ]
    },
    {
        "name": "Ohio Technical Institute",
        "website": "https://www.ohiotech.edu",
        "courses": [
            ("Complete Automotive Technology", "Automotive"),
            ("Diesel Equipment Technology", "Automotive"),
            ("Classic Car Restoration Technology", "Automotive"),
            ("Collision Repair and Refinishing", "Automotive"),
            ("High Performance and Racing / Alternative Fuel Vehicles", "Automotive"),
            ("PowerSport Technology", "Automotive"),
            ("Welding and Fabrication Technology (Master Welding Technology)", "Trades"),
            ("Power Generator Systems Technology", "Trades"),
            ("BMW STEP (Factory Advanced Skilled Training)", "Advanced/Specialty"),
            ("Custom Paint and Graphics", "Specialty"),
            ("Race Car Technology", "Specialty"),
        ]
    },
    {
        "name": "Pacific Institute of Culinary Arts",
        "website": "https://www.picachef.com",
        "courses": [
            ("Grand Diploma – Culinary and Baking & Pastry Arts (combined 6-month intensive)", "Diploma"),
            ("Grand Diploma + CO-OP (with paid one-year co-op work placement)", "Diploma"),
            ("Culinary Arts (6-month professional diploma)", "Diploma"),
            ("Pastry and Baking Arts (6-month professional diploma)", "Diploma"),
            ("Hospitality in Food and Beverage (F&B Management professional diploma)", "Diploma"),
            ("Casual Cooking Classes (recreational/public cooking classes – various topics)", "Recreational"),
        ]
    },
    {
        "name": "Porter and Chester Institute",
        "website": "https://porterchester.edu",
        "courses": [
            ("Dental Assisting", "Healthcare"),
            ("Medical Assisting", "Healthcare"),
            ("Practical Nursing (LPN)", "Healthcare"),
            ("HVACR (Heating, Ventilation, Air Conditioning, and Refrigeration)", "Trades"),
            ("Electrician: Industrial, Commercial and Residential", "Trades"),
            ("Low Voltage Technology", "Trades"),
            ("Plumbing", "Trades"),
            ("Welding", "Trades"),
            ("Automotive Technology", "Trades"),
            ("Computer Aided Drafting and Design (CADD)", "Technology"),
            ("Computer and Network Technology", "Technology"),
            ("Computer Aided Drafting and Design (CADD) – Online", "Technology - Online"),
            ("Computer and Network Technology – Online", "Technology - Online"),
        ]
    },
    {
        "name": "STVT-AAI Education",
        "website": "https://www.stvt.edu / https://www.aai.edu",
        "courses": [
            # STVT
            ("Medical Clinical Assistant (STVT)", "Diploma/Certificate"),
            ("Dental Assistant (STVT)", "Diploma/Certificate"),
            ("Medical Billing & Coding (STVT)", "Diploma/Certificate"),
            ("Automotive Technology (STVT)", "Diploma/Certificate"),
            ("Combination Welding (STVT)", "Diploma/Certificate"),
            ("HVAC/R – Heating, Ventilation, Air Conditioning & Refrigeration (STVT)", "Diploma/Certificate"),
            ("Diesel Heavy Truck Technician (STVT)", "Diploma/Certificate"),
            ("Electrical Installation (STVT)", "Diploma/Certificate"),
            ("CDL Training – Class A Tractor Trailer (STVT)", "Diploma/Certificate"),
            ("CDL Training – Class B (STVT)", "Diploma/Certificate"),
            ("Business Studies (STVT)", "Diploma/Certificate"),
            # AAI
            ("Automotive Service Technician (AAI – Arizona Automotive Institute)", "Diploma/Certificate"),
            ("Combination Welding (AAI)", "Diploma/Certificate"),
            ("Construction & Trades Management (AAI)", "Diploma/Certificate"),
            ("Diesel/Heavy Truck Technician (AAI)", "Diploma/Certificate"),
            ("Electrical (AAI)", "Diploma/Certificate"),
            ("HVAC & Basic Refrigeration (AAI)", "Diploma/Certificate"),
        ]
    },
    {
        "name": "Salon Success Academy",
        "website": "https://www.salonsuccessacademy.com",
        "courses": [
            ("Cosmetology", "Diploma/Certificate"),
            ("Esthetics (Esthetician)", "Diploma/Certificate"),
            ("Barbering", "Diploma/Certificate"),
            ("Nail Technology (Manicuring)", "Diploma/Certificate"),
            ("Make-Up Designory (MUD Makeup)", "Continuing Education"),
        ]
    },
    {
        "name": "Southeastern College",
        "website": "https://www.sec.edu",
        "notes": "Also known as Bar Education Inc. DBA Southeastern College",
        "courses": [
            ("Business Administration", "Diploma/Degree"),
            ("Cosmetology", "Diploma/Degree"),
            ("Diagnostic Medical Sonography", "Diploma/Degree"),
            ("Early Childhood Education", "Diploma/Degree"),
            ("HVAC (Heating, Ventilation, and Air Conditioning)", "Diploma/Degree"),
            ("Medical Assisting", "Diploma/Degree"),
            ("Nursing – Associate of Science / Associate of Applied Science", "Associate's Degree"),
            ("Practical Nursing (LPN)", "Diploma/Certificate"),
            ("RN to BSN (Bachelor of Science in Nursing)", "Bachelor's Degree"),
            ("Occupational Therapy Assistant", "Diploma/Degree"),
            ("Pharmacy Technology", "Diploma/Certificate"),
            ("Radiologic Technology", "Diploma/Degree"),
            ("Surgical Technology", "Diploma/Degree"),
            ("Medical Billing", "Diploma/Certificate"),
            ("Professional Clinical Massage Therapy", "Diploma/Certificate"),
        ]
    },
    {
        "name": "TWS Acquisition Corporation",
        "website": "https://www.tws.edu / https://www.refrigerationschool.com",
        "notes": "Operating as StrataTech Education Group – parent of Tulsa Welding School (TWS) and The Refrigeration School, Inc. (RSI)",
        "courses": [
            # TWS
            ("Professional Welder (TWS – Tulsa Welding School)", "Diploma/Certificate"),
            ("Welding Specialist with Pipefitting (TWS)", "Diploma/Certificate"),
            ("Associate of Occupational Studies (AOS) in Welding Technology (TWS)", "Associate's Degree"),
            ("Associate in Applied Science (AAS) in Welding Inspection & Quality Management (TWS)", "Associate's Degree"),
            ("Refrigeration Technologies / HVAC/R (TWS)", "Diploma/Certificate"),
            ("Electrical Applications (TWS)", "Diploma/Certificate"),
            ("Electro-Mechanical Technologies (TWS)", "Diploma/Certificate"),
            ("Electrical Lineworker (TWS)", "Diploma/Certificate"),
            ("Advanced Industrial Maintenance Technology (TWS)", "Diploma/Certificate"),
            # RSI
            ("HVAC/Refrigeration (RSI – The Refrigeration School)", "Diploma/Certificate"),
            ("Refrigeration Technologies (RSI)", "Diploma/Certificate"),
            ("Electrical Applications (RSI)", "Diploma/Certificate"),
            ("Professional Welder (RSI)", "Diploma/Certificate"),
            ("Mechanical Maintenance Engineering AOS (RSI)", "Associate's Degree"),
        ]
    },
    {
        "name": "Tricoci University of Beauty Culture",
        "website": "https://www.tricociuniversity.edu",
        "courses": [
            ("Cosmetology", "Diploma/Certificate"),
            ("Esthetics", "Diploma/Certificate"),
            ("Barbering", "Diploma/Certificate"),
            ("Nail Technology (Manicuring)", "Diploma/Certificate"),
            ("Cosmetology Teacher Training", "Diploma/Certificate"),
            ("Esthetics Teacher Training", "Diploma/Certificate"),
            ("Barber Teacher Training", "Diploma/Certificate"),
        ]
    },
    {
        "name": "UniteK Learning",
        "website": "https://www.uniteklearning.com / https://www.unitekcollege.edu / https://www.unitekemt.com",
        "courses": [
            # Unitek College
            ("Bachelor of Science in Nursing (BSN) (Unitek College)", "Bachelor's Degree"),
            ("LVN to BSN – Advanced Placement (Unitek College)", "Bachelor's Degree"),
            ("Vocational Nursing (LVN) (Unitek College)", "Diploma/Certificate"),
            ("Practical Nursing (LPN) – Nevada campuses (Unitek College)", "Diploma/Certificate"),
            ("Associate of Science in Vocational Nursing (ASVN) (Unitek College)", "Associate's Degree"),
            ("Medical Assisting (Unitek College)", "Diploma/Certificate"),
            ("Medical Office Administration (Unitek College)", "Diploma/Certificate"),
            ("Dental Assisting (Unitek College)", "Diploma/Certificate"),
            ("Physical Therapist Assistant (PTA) (Unitek College)", "Diploma/Certificate"),
            ("Information Technology (Unitek College)", "Diploma/Certificate"),
            # Unitek EMT
            ("EMT Boot Camp 14-Day (Unitek EMT)", "Certificate"),
            ("EMT Boot Camp 7-Week Evening (Unitek EMT)", "Certificate"),
            ("EMT Refresher Course (Unitek EMT)", "Continuing Education"),
            ("Psychomotor Skills Refresher (Unitek EMT)", "Continuing Education"),
            ("Basic EKG Course (Unitek EMT)", "Certificate"),
        ]
    },
]


def create_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "College Courses"

    # Styles
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=12)
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    college_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    college_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    alt_row_fill = PatternFill(start_color="EBF3FB", end_color="EBF3FB", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )

    # Headers
    headers = ["#", "College / Institute", "Official Website", "Course / Program Name", "Category / Level"]
    ws.append(headers)
    for col_idx, _ in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border

    # Set column widths
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 45
    ws.column_dimensions["D"].width = 65
    ws.column_dimensions["E"].width = 32

    row_num = 2
    for idx, college in enumerate(college_data, 1):
        name = college["name"]
        website = college["website"]
        courses = college["courses"]
        num_courses = len(courses)

        # College header row (merged across course rows)
        start_row = row_num

        for c_idx, (course_name, category) in enumerate(courses):
            ws.cell(row=row_num, column=1).value = idx
            ws.cell(row=row_num, column=1).alignment = center_align
            ws.cell(row=row_num, column=1).border = thin_border

            ws.cell(row=row_num, column=2).value = name
            ws.cell(row=row_num, column=2).alignment = left_align
            ws.cell(row=row_num, column=2).border = thin_border

            ws.cell(row=row_num, column=3).value = website
            ws.cell(row=row_num, column=3).alignment = left_align
            ws.cell(row=row_num, column=3).border = thin_border

            ws.cell(row=row_num, column=4).value = course_name
            ws.cell(row=row_num, column=4).alignment = left_align
            ws.cell(row=row_num, column=4).border = thin_border

            ws.cell(row=row_num, column=5).value = category
            ws.cell(row=row_num, column=5).alignment = left_align
            ws.cell(row=row_num, column=5).border = thin_border

            # Alternating row shading per college group
            if idx % 2 == 0:
                for col in range(1, 6):
                    ws.cell(row=row_num, column=col).fill = alt_row_fill

            row_num += 1

        # Style the college name/number/website columns
        for r in range(start_row, start_row + num_courses):
            ws.cell(row=r, column=2).font = Font(name="Calibri", bold=True, size=10)
            ws.cell(row=r, column=3).font = Font(name="Calibri", size=10, color="1F4E79", underline="single")

    # Freeze top row
    ws.freeze_panes = "A2"

    # Auto-filter
    ws.auto_filter.ref = f"A1:E{row_num - 1}"

    # Add a summary sheet
    ws_summary = wb.create_sheet(title="Summary")
    ws_summary.column_dimensions["A"].width = 5
    ws_summary.column_dimensions["B"].width = 42
    ws_summary.column_dimensions["C"].width = 48
    ws_summary.column_dimensions["D"].width = 18

    summary_headers = ["#", "College / Institute", "Official Website", "Total Programs"]
    ws_summary.append(summary_headers)
    for col_idx, _ in enumerate(summary_headers, 1):
        cell = ws_summary.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border

    for idx, college in enumerate(college_data, 1):
        ws_summary.append([
            idx,
            college["name"],
            college["website"],
            len(college["courses"])
        ])
        for col in range(1, 5):
            cell = ws_summary.cell(row=idx + 1, column=col)
            cell.alignment = left_align
            cell.border = thin_border
            if idx % 2 == 0:
                cell.fill = alt_row_fill
        ws_summary.cell(row=idx + 1, column=1).alignment = center_align
        ws_summary.cell(row=idx + 1, column=4).alignment = center_align

    ws_summary.freeze_panes = "A2"

    filename = "/home/user/ClaudeSkillsRepo/college_courses.xlsx"
    wb.save(filename)
    print(f"Excel file saved: {filename}")
    total_courses = sum(len(c["courses"]) for c in college_data)
    print(f"Total colleges: {len(college_data)}")
    print(f"Total courses/programs: {total_courses}")


if __name__ == "__main__":
    create_excel()
