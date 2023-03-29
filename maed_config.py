from pick import pick
def run(country):
    # maed_type = pick(["maedd", "maedel"], "Which MAED module do you want to use?", indicator='=>')
    maed_type="maedd"
    if maed_type == "maedd":
        selected_option, _ = pick(["Transport Starter Kits", "IEA"], "Which MAED template do you want to use?", indicator='=>')
        if selected_option == "IEA":
            scenario = ""
            maed_years = maed_years=[2022, 2025, 2030, 2035, 2040, 2045, 2050]
        else:
            scenario = input("Enter a scenario (BAU, MS, NZ): ")
            maed_years = input("Enter years for MAED analysis separated by commas: ")
            maed_years=maed_years.split(",")
            maed_years = [eval(i) for i in maed_years]
    elif maed_type == "maedel":
        selected_option = "MAED-EL"
        scenario = ""
        maed_years = [2020]
    return maed_type, selected_option, scenario, maed_years

if __name__ == "__main__":
    country = input("Enter a country: ")
    maed_type, selected_option, scenario, maed_years = run(country)
    print(maed_type)
    print(maed_years)