facilities_type = ["Top Facilities"]
categories = ["Beach"]
sub_categories = ["Private Beach", "Beach Front", "NA"]
availabilities = ["All year", "Seasonal", "NA"]
age_restrictions = ["All ages", "Kids Only", "Adults Only", "NA"]
cost_type = ["Paid", "Free", "NA"]

for facility in facilities_type:
    for category in categories:
        for sub_category in sub_categories:
            for availability in availabilities:
                for age_restriction in age_restrictions:
                    for cost in cost_type:
                        print(f"{facility};{category};{category};{sub_category};{age_restriction};{cost}")
