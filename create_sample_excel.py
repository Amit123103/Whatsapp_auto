import pandas as pd

# Generate sample leads with the exact enterprise format requested:
# #, Company, Industry, Lead Type, Country, City, Phone, Website, Website Status, LinkedIn (company), Google Rating, Source, Suggested Services, Contact Score, Priority, Outreach Status, Notes

sample_leads_data = [
    {
        "#": 1,
        "Company": "Apex Retail Solutions",
        "Industry": "Retail & E-Commerce",
        "Lead Type": "Cold Prospect",
        "Country": "India",
        "City": "Lucknow",
        "Phone": "9812345670",
        "Website": "https://apexretail.in",
        "Website Status": "Slow Loading (5.8s)",
        "LinkedIn (company)": "https://linkedin.com/company/apex-retail-sol",
        "Google Rating": "4.2 (85 reviews)",
        "Source": "Google Maps",
        "Suggested Services": "Modern Website Development & WhatsApp Store",
        "Contact Score": 92,
        "Priority": "High",
        "Outreach Status": "New",
        "Notes": "Your product catalog could convert 30% higher with a lightning-fast mobile website and direct WhatsApp checkout."
    },
    {
        "#": 2,
        "Company": "Zenith Healthcare Clinics",
        "Industry": "Healthcare & Diagnostics",
        "Lead Type": "Inbound Referral",
        "Country": "India",
        "City": "Delhi NCR",
        "Phone": "+919823456781",
        "Website": "https://zenithhealth.org",
        "Website Status": "Outdated Design",
        "LinkedIn (company)": "https://linkedin.com/company/zenith-healthcare",
        "Google Rating": "4.7 (194 reviews)",
        "Source": "Direct Referral",
        "Suggested Services": "Custom Mobile App (Android & iOS)",
        "Contact Score": 88,
        "Priority": "High",
        "Outreach Status": "New",
        "Notes": "An automated patient booking & digital prescription app would save your front-desk 2+ hours daily."
    },
    {
        "#": 3,
        "Company": "Mehta Logistics & Warehousing",
        "Industry": "Logistics & Transport",
        "Lead Type": "Cold Prospect",
        "Country": "India",
        "City": "Mumbai",
        "Phone": "9834567892.0",  # Test float format commonly exported by Excel
        "Website": "https://mehtalogistics.com",
        "Website Status": "No Tracking Portal",
        "LinkedIn (company)": "https://linkedin.com/company/mehta-logistics",
        "Google Rating": "4.1 (60 reviews)",
        "Source": "Justdial",
        "Suggested Services": "Custom ERP & Driver Tracking Software",
        "Contact Score": 85,
        "Priority": "Medium",
        "Outreach Status": "New",
        "Notes": "Automating driver dispatch and real-time consignment updates via WhatsApp will eliminate dispatch bottlenecks."
    },
    {
        "#": 4,
        "Company": "Gupta Fashion Studio",
        "Industry": "Fashion & Apparel",
        "Lead Type": "Local Business",
        "Country": "India",
        "City": "Jaipur",
        "Phone": "+91 98456 78903",  # Test spaces with +91
        "Website": "None",
        "Website Status": "No Website",
        "LinkedIn (company)": "None",
        "Google Rating": "4.5 (112 reviews)",
        "Source": "Instagram Leads",
        "Suggested Services": "Digital Marketing & Local SEO",
        "Contact Score": 79,
        "Priority": "Medium",
        "Outreach Status": "New",
        "Notes": "Ranking on Google Maps for local premium ethnic wear searches could double your weekend store footfall."
    },
    {
        "#": 5,
        "Company": "Singh Real Estate Promoters",
        "Industry": "Real Estate & Construction",
        "Lead Type": "Cold Prospect",
        "Country": "India",
        "City": "Bangalore",
        "Phone": "9856789014",
        "Website": "https://singhpromoters.com",
        "Website Status": "Needs Redesign",
        "LinkedIn (company)": "https://linkedin.com/company/singh-promoters",
        "Google Rating": "4.3 (78 reviews)",
        "Source": "Google Search",
        "Suggested Services": "High-Converting Landing Page & WhatsApp CRM",
        "Contact Score": 94,
        "Priority": "Urgent",
        "Outreach Status": "New",
        "Notes": "A dedicated property showcase landing page with automated WhatsApp lead capture will secure buyers before competitors do."
    },
    {
        "#": 6,
        "Company": "Apex Retail Solutions (Duplicate Entry)",
        "Industry": "Retail",
        "Lead Type": "Cold Prospect",
        "Country": "India",
        "City": "Lucknow",
        "Phone": "9812345670",  # Duplicate number to test duplicate prevention
        "Website": "https://apexretail.in",
        "Website Status": "Active",
        "LinkedIn (company)": "None",
        "Google Rating": "4.2 (85 reviews)",
        "Source": "Google Maps",
        "Suggested Services": "Website Development",
        "Contact Score": 70,
        "Priority": "Low",
        "Outreach Status": "New",
        "Notes": "Testing duplicate number detection filter."
    },
    {
        "#": 7,
        "Company": "Faulty Data Test Client",
        "Industry": "Technology",
        "Lead Type": "Invalid Test",
        "Country": "India",
        "City": "Delhi",
        "Phone": "12345",  # Invalid phone number to test automatic skip & continuation
        "Website": "None",
        "Website Status": "Inactive",
        "LinkedIn (company)": "None",
        "Google Rating": "3.5 (10 reviews)",
        "Source": "Manual Entry",
        "Suggested Services": "Cloud Migration",
        "Contact Score": 30,
        "Priority": "Low",
        "Outreach Status": "New",
        "Notes": "Testing invalid phone number handling."
    }
]

df = pd.DataFrame(sample_leads_data)
output_file = "sample_leads.xlsx"
df.to_excel(output_file, index=False)
print(f"Successfully generated '{output_file}' with {len(df)} rows and {len(df.columns)} columns!")
