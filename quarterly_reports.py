import os
import requests
from fpdf import FPDF
from fpdf import Template
import math
import configparser
import sys


def field_or_default(arr, field, default=None):
    if field in arr:
        return arr[field]
    else:
        return default

# Return records in requested table using Airtable API


def get_table_from_airtable(table_name):
    offset = None
    airtable_records = []
    while True:
        base_id = config['DEFAULT']['BaseID']
        url = "https://api.airtable.com/v0/" + base_id + "/" + table_name
        params = {}
        if offset is not None:
            params = {'offset': offset}
        api_key = config['DEFAULT']['APIKey']
        headers = {"Authorization": "Bearer " + api_key}
        response = requests.get(url, params=params, headers=headers)
        airtable_response = response.json()
        airtable_records.extend(airtable_response['records'])
        if 'offset' in airtable_response:
            offset = airtable_response['offset']
        else:
            break
    return airtable_records


if __name__ == '__main__':
    config = configparser.ConfigParser()
    config.read('config.ini')

    quarter = sys.argv[1]
    year = sys.argv[2]
    quarters = {
        'Q1': '03-31',
        'Q2': '06-30',
        'Q3': '09-30',
        'Q4': '12-31'
    }

    date = year + "-" + quarters[quarter]

    # Pulls Company Airtable and creates a id<>name lookup table
    companies = get_table_from_airtable("Companies")
    company_lookup = {}
    for record in companies:
        fields = record['fields']
        if fields['Status'] != "Scout":
            company_lookup[record['id']] = {
                'Name': fields['Name'],
                'Location': field_or_default(fields, 'Location'),
                'CEO': field_or_default(fields, 'CEO'),
                'Vehicles': field_or_default(fields, 'Vehicles'),
                'Company Description': field_or_default(fields, 'Company Description'),
                'Quarterly Update': field_or_default(fields, 'Quarterly Update'),
                'Logo': field_or_default(fields, 'Logo'),
                'URL': field_or_default(fields, 'URL'),
                'Initial Investment': field_or_default(fields, 'Initial Investment'),
                'Status': field_or_default(fields, 'Status'),
                'Valuation': field_or_default(fields, 'Valuation')}

    # Pulls Vehicle Airtable and creates a id<>name lookup table
    vehicles = get_table_from_airtable("Vehicles")
    vehicle_lookup = {}
    for record in vehicles:
        fields = record['fields']
        if "Fund" in record['fields']['Name']:
            vehicle_lookup[record['id']] = {
                'Name': field_or_default(fields, 'Name'), 'Logo': field_or_default(fields, 'Logo')}

    # Pulls Founder Airtable and creates id<>name lookup table
    founders = get_table_from_airtable("Founders")
    founder_lookup = {}
    for record in founders:
        fields = record['fields']
        founder_lookup[record['id']] = {
            'Full Name': field_or_default(fields, 'Full Name')}

    # Pulls investment rounds table from Airtable
    investment_rounds = get_table_from_airtable("Investment Rounds")

    # Generate summary table that will later feed into quarterly report table
    summaries = []
    for round in investment_rounds:
        fields = round['fields']
        if fields['Company'][0] in company_lookup:
            summary = {'Company': company_lookup[fields['Company'][0]]['Name'], 'Investment Round': field_or_default(fields, 'Investment Round'),
                       'Vehicle': field_or_default(fields, 'Vehicle'), 'Date': field_or_default(fields, 'Date'),
                       'Round Size': field_or_default(fields, 'Round Size'),
                       'Entry Valuation': field_or_default(fields, 'Entry Valuation (Post or Cap)'),
                       'Root Investment': field_or_default(fields, 'Root Investment Cost'),
                       'Total Value': field_or_default(fields, 'Total Value'),
                       'Root FD % (Last Closed Round)': field_or_default(fields, 'Root FD % (Last Closed Round)')}
            summaries.append(summary)
    summaries = sorted(summaries, key=lambda x: x['Date'])

    # Create a new PDF
    pdf = FPDF(format='Letter')

    # Iterate through each fund vehicle (in order Fund I --> III) and create title page + quarterly updates
    for vehicle in sorted(vehicle_lookup, key=lambda k: vehicle_lookup[k]['Name']):
        vehicle_name = vehicle_lookup[vehicle]['Name']

        # Create vehicle cover page with Root Logo
        quarter_header = quarter + " " + year
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        pdf.image(vehicle_lookup[vehicle]['Logo'][0]['url'],
                  x=87.95, y=100, w=40, type='', link='')
        pdf.ln(h='1')
        pdf.set_y(150)
        pdf.cell(0, 10, txt=vehicle_name, ln=1, align="C")
        pdf.cell(0, 7, txt="Operational Summaries", ln=1, align="C")
        pdf.cell(0, 7, txt=quarter_header, ln=1, align="C")

        # Iterate through alphabetically sorted companies within the specific vehicle and generate the pages of operational reports

        for company in sorted(company_lookup, key=lambda k: company_lookup[k]['Valuation'], reverse=True):
            valuation = company_lookup[company]['Valuation']

            if (vehicle in company_lookup[company]['Vehicles']) and (company_lookup[company]['Initial Investment'][0] <= date) and (company_lookup[company]['Status'] == 'Active'):
                logo_url = company_lookup[company]['Logo'][0]['url'] or ""
                if company_lookup[company]['CEO'] is not None:
                    founder_txt = "CEO: " + \
                        (founder_lookup[company_lookup[company]
                         ['CEO'][0]]['Full Name'] or "")
                else:
                    founder_txt = "CEO: "
                location_txt = "Location: " + \
                    (company_lookup[company]['Location'] or "")
                website_txt = "Website: " + \
                    (company_lookup[company]['URL'] or "")

                fd_ownership = 0
                for summary in summaries:
                    if (summary['Company'] == company_lookup[company]['Name']) and (summary['Date'] <= date) and (summary['Vehicle'][0] in vehicle_lookup):
                        if (type(summary['Root FD % (Last Closed Round)']) is dict) or (summary['Root FD % (Last Closed Round)'] is None):
                            fd_ownership = None
                            break
                        else:
                            fd_ownership = fd_ownership + \
                                summary['Root FD % (Last Closed Round)']
                if fd_ownership == None:
                    ownership_txt = "Ownership: N/A"
                else:
                    ownership_txt = "Ownership: {:.2%}".format(fd_ownership)
                description_txt = company_lookup[company]['Company Description'] or ""
                description_txt2 = description_txt.encode(
                    'latin-1', 'replace').decode('latin-1')

                # Add page and start the formatting
                pdf.add_page()
                pdf.set_font("Arial", size=12)

                # Title/Overview
                pdf.image(logo_url, x=77.95, y=None, w=60, type='', link='')
                pdf.set_font('Arial', 'B', 12)
                pdf.set_y(55)
                pdf.cell(
                    0, 7, txt='Overview', ln=1, align="L")
                pdf.set_font('')
                pdf.cell(
                    0, 7, txt=founder_txt, ln=0, align="L")
                pdf.cell(
                    0, 7, txt=location_txt, ln=1, align="R")
                pdf.cell(0, 7, txt=website_txt, ln=0, align="L")
                pdf.cell(0, 7, txt=ownership_txt, ln=1, align="R")

                # Description
                pdf.ln(h='1')
                pdf.multi_cell(
                    0, 7, txt=description_txt2, align="L")

                # Financing Table Setup
                pdf.ln(h='1')
                pdf.set_font('Arial', 'B', 12)
                pdf.cell(
                    0, 7, txt="Financing (Unaudited)", ln=2, align="L")
                pdf.set_font('')
                pdf.set_font('Arial', 'B', 12)
                pdf.set_fill_color(200., 200., 200.)
                pdf.cell(
                    52, 7, txt="Investment Round", border=1, ln=0, align="L", fill=1)
                pdf.cell(
                    30, 7, txt="Date", border=1, ln=0, align="L", fill=1)
                pdf.cell(
                    30, 7, txt="Round Size", border=1, ln=0, align="L", fill=1)
                pdf.cell(
                    30, 7, txt="Post or Cap", border=1, ln=0, align="L", fill=1)
                pdf.cell(
                    27, 7, txt="Invested", border=1, ln=0, align="L", fill=1)
                pdf.cell(
                    27, 7, txt="Fair Value", border=1, ln=1, align="L", fill=1)
                pdf.set_font('')

                # Here's where you print the content of the table
                total_invested = 0.
                total_fair_value = 0.
                for summary in summaries:
                    if (summary['Company'] == company_lookup[company]['Name']) and (summary['Date'] <= date) and (summary['Vehicle'][0] in vehicle_lookup):
                        total_invested += float(
                            summary['Root Investment'] or 0)
                        total_fair_value += float(summary['Total Value'] or 0)

                        pdf.cell(
                            52, 7, txt=summary['Investment Round'] or "??", border=1, ln=0, align="L")
                        pdf.cell(
                            30, 7, txt=summary['Date'] or "??", border=1, ln=0, align="L")
                        pdf.cell(
                            30, 7, txt='${:,.0f}'.format(float(summary['Round Size'] or 0)), border=1, ln=0, align="L")
                        pdf.cell(
                            30, 7, txt='${:,.0f}'.format(float(summary['Entry Valuation'] or 0)), border=1, ln=0, align="L")
                        pdf.cell(
                            27, 7, txt='${:,.0f}'.format(float(summary['Root Investment'] or 0)), border=1, ln=0, align="L")
                        pdf.cell(
                            27, 7, txt='${:,.0f}'.format(float(summary['Total Value'] or 0)), border=1, ln=1, align="L")

                # Total row at bottom of table
                pdf.set_font('Arial', 'B', 12)
                pdf.cell(
                    52, 7, txt='', border=0, ln=0, align="L")
                pdf.cell(
                    30, 7, txt='', border=0, ln=0, align="L")
                pdf.cell(
                    30, 7, txt='', border=0, ln=0, align="L")
                pdf.cell(
                    30, 7, txt='Total: ', border=0, ln=0, align="R")
                pdf.cell(
                    27, 7, txt='${:,.0f}'.format(total_invested), border=1, ln=0, align="L")
                pdf.cell(
                    27, 7, txt='${:,.0f}'.format(total_fair_value), border=1, ln=1, align="L")

                # Add the operational update here
                pdf.ln(h='1')
                pdf.set_font('Arial', 'B', 12)
                pdf.cell(
                    0, 7, txt='Operational Update', ln=1, align="L")
                pdf.set_font('')
                update_txt = company_lookup[company]['Quarterly Update'] or ""
                update_txt2 = update_txt.encode(
                    'latin-1', 'replace').decode('latin-1')
                pdf.multi_cell(
                    0, 7, txt=update_txt2, align="L")

    file_name = year + "-" + quarter + ".pdf"
    pdf.output(file_name)
