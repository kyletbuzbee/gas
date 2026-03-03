import json

def create_blueprint():
    blueprint = {
        "entities": {
            "Company": {
                "primary_key": "Company ID",
                "canonical_name_source": "Prospects.csv",
                "fields": {
                    "Company ID": "Unique Identifier (CID-XXXX)",
                    "Company Name": "Canonical Name (Title Case)",
                    "Industry": "Categorization from Settings.csv",
                    "Address": "Site Location",
                    "Zip Code": "Postal Code",
                    "Priority Score": "Calculated value for sales focus"
                },
                "relationships": {
                    "Contacts": "One-to-Many via Company ID",
                    "Outreach": "One-to-Many via Company ID",
                    "Transactions": "One-to-Many via Company ID"
                }
            },
            "Contact": {
                "primary_key": "Email",
                "fields": {
                    "Name": "Full Name",
                    "Company ID": "Link to Company",
                    "Role": "Contact Role",
                    "Phone": "Standardized (XXX) XXX-XXXX"
                }
            },
            "Outreach": {
                "primary_key": "Outreach ID",
                "fields": {
                    "Outreach ID": "Unique Identifier (LID-XXXXX)",
                    "Company ID": "Link to Company",
                    "Visit Date": "Standardized Date",
                    "Outcome": "Standardized from Workflow Rules",
                    "Notes": "Rich text outreach logs"
                }
            },
            "Transaction": {
                "primary_key": "Transaction ID",
                "fields": {
                    "Transaction ID": "Unique Identifier (TID-XXXXX)",
                    "Company ID": "Link to Company",
                    "Date": "Standardized Date",
                    "Material": "Standardized from Settings.csv",
                    "Net Weight": "Numeric weight",
                    "Price": "Numeric rate",
                    "Total Payment": "Calculated (Weight * Price)"
                }
            }
        },
        "normalization_rules": {
            "dates": "MM/DD/YYYY",
            "phones": "(999) 999-9999",
            "currency": "Float (no symbols)",
            "names": "Title Case (Stripped)"
        }
    }

    with open("crm_custom_blueprint.json", "w") as f:
        json.dump(blueprint, f, indent=4)
    print("CRM Blueprint generated: crm_custom_blueprint.json")

if __name__ == "__main__":
    create_blueprint()