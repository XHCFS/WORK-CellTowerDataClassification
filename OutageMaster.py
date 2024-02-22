import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import argparse
import re
from datetime import datetime, timedelta
from fuzzywuzzy import fuzz
import os

officeName="C:\\Users\\swx1283483\\automation-scripts\\Office(1).xlsx"
corpName="C:\\Users\\swx1283483\\automation-scripts\\Corpp.xlsx"

custom_rules = {
    r"(HW|SW)": {
        "Reason Category": "BSS",
        "Reason Sub-Category": "BSS_HW"
    },

    r"SW": {
        "Reason Category": "BSS",
        "Reason Sub-Category": "BSS_SW"
    },

    r"HT": {
        "Reason Category": "High_temp",
        "Reason Sub-Category": "High_temp"
    },


    r"TX": {
        "Reason Category": "TX",
        "Reason Sub-Category": ""
    }


#    r"": {
#        "Reason Category": "",
#        "Reason Sub-Category": ""
#    }
#


    # Add more rules here...
}



desired_header_order = ["Date", "Region", "BSC", "Site Name", "Site ID", "2G/3G", "ID","Site Category","Priority","Technical Area","Site Layer - Qism","Type Of Sharing","Host/Guest","Alarm Occurance Time","Fault Occurance Time","Fault Clearance Time","MTTR","Hybrid Down Time","SLA Status","Site Type", "Reason Category","Reason Sub-Category","Comment", "Owner","Access Type","Cascaded To", "Final", "Access Category", "Office","Corp","Generator","Vendor","Reg", "Most Aff"]



desired_header_order_mt= ["Site", "TECH", "ID", "Controller", "RBSType", "Region", "nm_tier","EventTime","CeaseTime","MTTR","Duration","Reason Category","Reason Sub-Category","RootCause","DateOfOutage","Vendor","FilledBy","MR_CP","Office","Site Type"]


original_headers= {"Reason Category":["BSS", "High_temp", "Others", "Power", "TX",""],

                       "Reason Sub-Category":['BSS_HW','BSS_SW','High_Temp','High_Temp_Dependency','Others_Illegal_Intervention','Others_ROT','Others_Dependency_Illegal Intervention','Others_Unknown Reason','Others_Dependency_Unknown Reason','Others_Wrong_Action','Power_Commercial','Power_Dependency_Commercial','Power_Generator','Power_Dependency_Generator','Power_HW (Cct Breakers, Cables)','Power_Dependency_HW (Cct Breakers, Cables)','Power_Power Criteria','Power_Dependency_Power Criteria','Power_Solar Cell','Power_Dependency_Solar Cell','TX_Bad Performance','TX_Dependency_Bad Performance','TX_HW Failure','TX_Dependency_HW Failure','TX_LOS','TX_Dependency_LOS','TX_Physical Connection','TX_Dependency_Physical Connection','TX_Power Supply','TX_Dependency_Power Supply','TX_SW Failure','TX_Telecom Egypt','TX_Dependency_Telecom Egypt','BSS/License',""],

                       "Owner":["BSS","FM","GD","ROT","RT","TD","TX","E///","ZTE",""]
}


def merge_categories(original_categories, new_categories):
    # Initialize a dictionary to store mappings
    mappings = {}
    
    # Remove empty strings and non-string objects from input lists
    original_categories = [x for x in original_categories if isinstance(x, str) and x]
    new_categories = [x for x in new_categories if isinstance(x, str) and x]

    # Iterate through each new category
    for new_cat in new_categories:
        best_similarity = 0
        best_orig_cat = None
        
        # Iterate through each original category
        for orig_cat in original_categories:
            # Calculate the similarity score between the two categories
            similarity_score = fuzz.ratio(new_cat.lower(), orig_cat.lower())

            # If the similarity score is above a threshold (e.g., 80) and higher than the current best,
            # consider them a match
            if similarity_score >= 70 and similarity_score > best_similarity:
                best_similarity = similarity_score
                best_orig_cat = orig_cat
        
        # If a match is found, add it to the mapping
        if best_orig_cat:
            mappings[new_cat] = best_orig_cat
    mappings={key: value for key,value in mappings.items() if key!=value and key!=""}
    return mappings


def apply_cascaded_to_rule(row):
    comment = row["Comment"]
    if isinstance(comment, str):
        cas_matches = re.findall(r'cas', comment, re.IGNORECASE)
        
        if cas_matches and row["Reason Category"]!="Others":
            match = re.search(r'\d{3,4}', comment)
            if match:
                site_id_prefix = row["Site ID"][:3]
                site_id_postfix = row["Site ID"][-4:]
                cascaded_to_digits = match.group().zfill(4)

                if site_id_postfix == cascaded_to_digits:
                    row["Cascaded To"] = ""
                    pattern = r"(?i)(cascaded|cas|dependency|to|\b\d{3,4}\b)"
                    row['Comment']= re.sub(pattern, "", row['Comment'])
                    row["Reason Sub-Category"]=row["Reason Sub-Category"].replace("Dependancy","")

                else:
                    row["Cascaded To"] = site_id_prefix + cascaded_to_digits

                print(f"Modified Cascaded To in row {row.name}: '{row['Cascaded To']}'")
    return row


def apply_custom_rules(row, custom_rules):
    comment = row["Comment"]
    if isinstance(comment, str):
        for keyword, rule in custom_rules.items():
            # Compile the regular expression pattern from the keyword
            pattern = re.compile(keyword, re.IGNORECASE)  # Case-insensitive matching
            if pattern.search(comment):
                if rule["Reason Category"] != "":
                    row["Reason Category"] = rule["Reason Category"]

                if rule["Reason Sub-Category"] != "":
                    row["Reason Sub-Category"] = rule["Reason Sub-Category"]

                # Add a print statement here to describe the modification
                print(f"Modified Reason Category in row {row.name}: '{row['Reason Category']}'")
                print(f"Modified Reason Sub-Category in row {row.name}: '{row['Reason Sub-Category']}'")
                break  # Once a match is found, we apply the rule and break out of the loop
    return row

def apply_tcr_logic(row):
    comment = row["Comment"]
    if isinstance(comment, str):
        tcr_match = re.search(r'tcr', comment, re.IGNORECASE)

        if pd.notna(row["SLA Status"]):  # Check if "SLA Status" cell is not empty
            if tcr_match:
                sla_status = row["SLA Status"]
                if isinstance(sla_status, str):
                    sla_status_words = sla_status.split()
                    if sla_status_words and sla_status_words[0].lower() != "planned":
                        row["SLA Status"] = "Planned " + " ".join(sla_status_words[1:])
                        print(f"Modified SLA Status in row {row.name}: '{row['SLA Status']}'")
    return row

def remove_dependency_from_subcategory(row):
    comment = row["Comment"]
    sub_category = row["Reason Sub-Category"]

    if isinstance(comment, str) and isinstance(sub_category, str):
        if not re.search(r'cas', comment, re.IGNORECASE):
            # Remove "Dependency" from the sub-category using regex
            row["Reason Sub-Category"] = re.sub(r'_?Dependency$', '', sub_category, flags=re.IGNORECASE)
    
        elif re.search(r'cas', comment, re.IGNORECASE) and not re.search(r'_?Dependency', sub_category, re.IGNORECASE):
            # Add "Dependency" to the sub-category using regex
            row["Reason Sub-Category"] = f"{sub_category}_Dependency"
    
    return row




def fixes(data, custom_rules):
    data = data.apply(apply_cascaded_to_rule, axis=1)
    print(data['Cascaded To']) 
#   data = data.apply(lambda row: apply_custom_rules(row, custom_rules), axis=1)
    data = data.apply(apply_tcr_logic, axis=1)
    data = data.apply(remove_dependency_from_subcategory, axis=1)
    return data

def MostAff(row):
    comment = row["Comment"]
    if isinstance(comment, str):
        if isinstance(row['Cascaded To'], str):
            row['Most Aff']=row['Cascaded To']
        else:
            row['Most Aff']=row['Site ID']
        

        print(f"Modified Most Aff in row {row.name}: '{row['Most Aff']}'")
    else:
        row['Most Aff']=row['Site ID']
    return row

def DropEmptyComments(df):
    df=df[df['Comment'].apply(lambda x: isinstance(x, str))]    

    return df

def Reg(row):
    comment = row["Comment"]
    if isinstance(comment, str):
        row['Reg']=row['Site ID'][:3]
        print(f"Modified Reg in row {row.name}: '{row['Reg']}'")
    return row

def Generator(row):
    comment = row["Comment"]
    if isinstance(comment, str):
        if re.search(r'COMM|CRISIS|(?:^|\s)EC', comment, re.IGNORECASE):
            if re.search(r'CAS', comment, re.IGNORECASE):
                row["Reason Sub-Category"]='Power_Dependency_Commercial'
            else:
                row["Reason Sub-Category"]='Power_Commercial'

        if row['Reason Sub-Category'] in ['Power_Generator','Power_Dependency_Generator','Others_Illegal_Intervention'] :
            rules=[
            ["Planned",r"TCR|PLAN"],
            ["Stolen Fuel",r"STOLEN\sF|FUEL\sSTOLEN|F(?:\w*)\sSTOLEN"],
            ["Rental",r"REN"],
            ["Sharing",r"Shar"],
            ["Spare Part",r"SP"],
            ["Overhauling",r"OVERH|OH"],
            ["Fuel Problem",r"SHORTAGE"],
            ["Warranty ",r"war"],
            ["Tech Problem",r"RATOR|REN|SOLAR|ACCESS|OUTAGE|ATS|TNT|PMDG|GEN|DGMAIN"]
            ]

            for expression in rules:
                if re.search(expression[1], comment, re.IGNORECASE):
                    row["Generator"]=expression[0]
                    break

            print(f"Modified Generator in row {row.name}: '{row['Generator']}'")
    return row


def Final(row):
    comment = row["Comment"]
    siteName = row["Site Name"]
    if isinstance(comment, str):
        rulesComment=[
        ["Arish",r"ARISH"],
        ["Others",r"UNKNOWN"],
        ["Damaged/Stolen",r"DAMAGED|BURNT|INCIDENT"],
        ["Planned",r"TCR|PLANNED"],
        ["Power",r"SABOT|HCR|ODU|ATS|FULE|FEUL|FUEL|HYBRID|PHASE|(?:^|\s)EC|DG|RECT|HAYA|C\.?B|COMMERCIAL|COM|COMM|TRANS\S|TRANSFO|GENERATOR|SOLAR|POWER|GEN|RATOR|OUTAGE"],
        ["Access",r"ACES|ACCES|CONT|SHARING|OWNER"],
        ["High_Temp",r"TEMP"],
        ["BSS/TX",r"FTTS|ELR|RTN|PERF|FLAP|INTER|BSS|TX|LINK|CARD|FIBER|CABLE|SW|CONNECTION|HW|E1|SW"],
        ["Under ROT",r"ROT(?:\s|$)|TD|TCR"]
        ]

    rulesOffice=[["Arish",r"(?:\W|_)\WArish|(?:\W|_)NSN(?:\W|_)|(?:\W|_)ARS(?:\W|_)"]]
    for expression in rulesComment:
        if re.search(expression[1], comment, re.IGNORECASE):
            row["Final"]=expression[0]
            break
    for expression in rulesOffice:
        if re.search(expression[1], siteName, re.IGNORECASE):
            row["Final"]=expression[0]
            break

    print(f"Modified Generator in row {row.name}: '{row['Final']}'")
    return row


def Office(row,officeDict):
    try:
        row["Office"]= officeDict[row["Site ID"]]
    except Exception as e:
        print(e)
    return row



def Access_Category(row):
    comment = row["Comment"]
    final= row["Final"]
    if isinstance(comment, str):
        rulesComment=[
        ["Access_Guard Others",r"Guard"],
        ["Access_GD",r"GD|Railway|Police|(?:\s|^|\()TE(?:\s|$|\))"],
        ["Access_Owner Fee",r"Cheque|Payment|Cheqe"],
        ["End_Of_Contract",r"EOC"],
        ["Access_Owner Others",r"ARMY|SHARING|Owner|H&S|Milit|Hotel|Factory|"],
        ]


    for expression in rulesComment:
        if re.search(expression[1], comment, re.IGNORECASE) and final not in ["Power","BSS/TX"]:
            row["Access Category"]=expression[0]
            break

    print(f"Modified Access Category in row {row.name}: '{row['Access Category']}'")
    return row

def process_excel_files(input_files,input_files2 ,output_file):
    combined_data = pd.DataFrame(columns=desired_header_order)  
    for input_file in input_files:
        data = pd.read_excel(input_file)
        for header in desired_header_order:
            if header not in data.columns:
                data[header]=None

        data["Site ID"] = data["SiteCode"]
        data["2G/3G"] = data["Tech"]
        data["Site Layer - Qism"] = data["Site Layer Qism"]
        data["Hybrid Down Time"] = data["Down Time"].apply(lambda x: float(x)/60)
         

        data = data[desired_header_order]


        combined_data = pd.concat([combined_data, data])








    combined_data_mt = pd.DataFrame(columns=desired_header_order_mt)  # Initialize with desired order
    for input_file2 in input_files2:
        data2 = pd.read_excel(input_file2)
        for header in desired_header_order_mt:
            if header not in data.columns:
                data[header]=None

        data2["Site Type"] = data2["RBSType"] + " " + data2["ID/OD"]
        data2["ID"] = data2["Site"].str[-4:]
        data2["TECH"] = data2["Site"].apply(lambda x: "4G" if x[0] == "L" else "3G" if x[0] == "U" else "2G")
        data2["Site"] = data2["Site"].str[-7:]
        data2["Reason Category"] = data2["Category"]
        data2["Reason Sub-Category"] = data2["SubCategory"]

        for header in desired_header_order:
            if header not in data2.columns:
                data2[header]=None
        
         

        data2["Date"] = (datetime.now()- timedelta(days=1))
        data2["BSC"]=data2["Controller"]
        data2["2G/3G"]=data2["TECH"]

        data2["Priority"]=data2["nm_tier"]
        data2["Site ID"]=data2["Site"]
        data2["Site Name"]=data2["Site"]

        data2["Technical Area"]=data2["Region"]
        data2["Site Layer - Qism"]=data2["Region"]

        data2["Fault Clearance Time"]=data2["CeaseTime"]
        data2["Alarm Occurance Time"]=data2["EventTime"]
        data2["Fault Occurance Time"]=data2["EventTime"]
        data2["Hybrid Down Time"]=data2["Duration"].apply(lambda x: float(x)/60)
        data2["Comment"]=data2["RootCause"]
        data2 = data2[desired_header_order]

        combined_data_mt = pd.concat([combined_data_mt, data2])



    print(f"{combined_data.shape}{combined_data_mt.shape}")
    combined_data = pd.concat([combined_data, combined_data_mt])
    try:
        combined_data["Date"]=combined_data["Date"].apply(lambda x: x.strftime("%m-%d-%Y"))
    except:
        pass
    combined_data=combined_data[desired_header_order]

    for key,header in original_headers.items():
        mcd=merge_categories(header,combined_data[key].unique().tolist())
        print(combined_data[key].unique().tolist())
        print(f"\n{key}") 
        for og,nw in mcd.items():
            print(f"{og}:{nw}")
            combined_data[key] = combined_data[key].str.replace(og,nw)
    combined_data=DropEmptyComments(combined_data)
    combined_data=fixes(combined_data,custom_rules)
    combined_data= combined_data.apply(Final, axis=1)
    combined_data= combined_data.apply(Generator, axis=1)
    combined_data= combined_data.apply(Reg, axis=1)
    combined_data= combined_data.apply(Access_Category, axis=1)

    for key,header in original_headers.items():
        mcd=merge_categories(header,combined_data[key].unique().tolist())
        print(combined_data[key].unique().tolist())
        print(f"\n{key}") 
        for og,nw in mcd.items():
            print(f"{og}:{nw}")
            combined_data[key] = combined_data[key].str.replace(og,nw)


    combined_data= combined_data.apply(MostAff, axis=1)





    
    corp=pd.read_excel(corpName)
    corpList=corp["Facing Site"].tolist()
    office=pd.read_excel(officeName)
    officeDict= dict(zip(office["SiteCode"],office["Office"]))

    combined_data= combined_data.apply(Office,officeDict=officeDict ,axis=1)
    combined_data["Corp"]=combined_data["Site ID"].apply(lambda x: "Yes" if x in corpList else "No")




    combined_data.to_excel(output_file, index=False)

def get_files_in_current_directory():
    try:
        current_directory = os.getcwd()
        files = os.listdir(current_directory)
        return [x for x in files if x.endswith(".xlsx") and "Full" not in x]
    except Exception as e:
        print(f"Error getting files in current directory: {str(e)}")
        return []

def rename_files(input_files):
    i=0
    for input_file in input_files:
        if not os.path.exists(input_file):
            continue
        
        try:
            data = pd.read_excel(input_file)
            region = data["Region"].iloc[0].title()  # Capitalize the region name
             
            # Check if region is Menoufia or Tanta (case-insensitive)
            if region.lower() in ["menoufia", "tanta"]:
                date = (datetime.now()-timedelta(days=1)).strftime("%m-%d-%Y")
                i+=1
                # Apply secondary renaming rules for Menoufia or Tanta
                region = region.title()  # Ensure region name is capitalized
                new_file_name = f"Menoufia and Tanta {i} Outage Summary Report {date}.xlsx"
            else:
                date = data["Date"].iloc[0]  # Capitalize the region name
                tech = data["Tech"].iloc[0]
                # Check if tech is 4G, 3G, or 2G (case-insensitive)
                if tech.lower() in ["4g", "3g", "2g"]:
                    tech = tech.upper()  # Ensure tech is uppercase
                else:
                    tech = ""  # Empty string for other tech values
                
                # Apply original renaming rules for non-4G/3G/2G cases
                # Rename the file as "{tech} {region} Outage Summary Report" + date
                new_file_name = f"{tech} {region} Outage Summary Report {date.strftime('%m-%d-%Y')}.xlsx"
                
            # Rename the file
            os.rename(input_file, new_file_name)
            print(f"Renamed '{input_file}' to '{new_file_name}'")
        
        except Exception as e:
            print(f"Error renaming '{input_file}': {str(e)}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process multiple Excel files.")
    parser.add_argument("-i", "--input1", nargs="+", help="List of input Excel files")
    parser.add_argument("-m", "--input2", nargs="+", help="List of input Excel files")
    parser.add_argument("-o", "--output", help="Output Excel file")
    args = parser.parse_args()


    if not args.input1 and not args.input2:
        rename_files(get_files_in_current_directory())
        files=get_files_in_current_directory()
        input1 = [file for file in files if file.lower().startswith(("4g", "3g", "2g"))]
        input2 = [file for file in files if not file.lower().startswith(("4g", "3g", "2g"))]

        output=f'Full Outage Report Update {(datetime.now()-timedelta(days=1)).strftime("%m-%d-%Y")}.xlsx'
        process_excel_files(input1, input2, output)
    else:
        process_excel_files(args.input1, args.input2, args.output)
        rename_files(args.input1 + args.input2)
