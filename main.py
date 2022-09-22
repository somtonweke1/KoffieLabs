import pandas as pd

# This reads the Excel file and returns a pandas dataframe
def read_file() -> pd.DataFrame:
    df = pd.read_excel('input_data.xlsx')
    return df


# Given a state, this returns the tax rate for that state.
def get_tax_rate(state):

    state_rates = {'IL': .0251, 'TN': .01766, }
    return state_rates[state]


# This  takes in a dataframe and a date, and writes a new excel file with the aggregated data
def write_new_file(aggregated_data_frame: pd.DataFrame, report_date: str):

    aggregated_data_frame.to_excel(f'aggregated_report-{report_date}.xlsx')
    return

 
# This takes a dataframe and a report date as inputs, and returns a dataframe with the following columns added:
# Pro Rata GWP, Earned Premium, Unearned Premium, and Taxes.
def add_calculated_columns(dataframe, report_date):
    import datetime
    report_date_converted = datetime.datetime.strptime(report_date, "%Y-%m-%d")
    
    effective_dates = dataframe['Effective Date'].tolist()
    expiration_dates = dataframe['Expiration Date'].tolist()
    states = dataframe['State'].tolist()
    annual_gwps = dataframe['Annual GWP'].tolist()
    pro_rata_gwps = []
    earned_premiums = []
    unearned_premiums = []
    taxes = []

# This iterates through the states list.
    for i in range(0, len(states)):
        # 2021-10-O1
        effective_date = effective_dates[i]
        expiration_date = expiration_dates[i]
        annual_gwp = annual_gwps[i]
        
        # It calculates the pro rata premium, earned premium, unearned premium, and taxes for each policy.
        if isinstance(effective_date, datetime.datetime) and isinstance(expiration_date, datetime.datetime) and type(annual_gwp) in [int, float]:
            year = effective_date.year
            import calendar
            days_in_year = 365 + calendar.isleap(year)
            days_between = (expiration_date - effective_date).days
            daily_gwp = float(annual_gwp) / float(days_in_year)
            # pro rata gwp
            pro_rata_gwp = daily_gwp * days_between
            pro_rata_gwps.append(pro_rata_gwp)
            # earned and unearned premium
            days_between_effective_and_report = abs((report_date_converted - effective_date).days)
            days_between_expiration_and_report = abs((report_date_converted - expiration_date).days)
            earned_premium = days_between_effective_and_report * daily_gwp
            unearned_premium = days_between_expiration_and_report * daily_gwp
            earned_premiums.append(earned_premium)
            unearned_premiums.append(unearned_premium)
            # taxes
            rate = get_tax_rate(states[i])
            taxes.append(earned_premium * rate)

        else:
            # invalid data
            pro_rata_gwps.append(0)
            earned_premiums.append(0)
            unearned_premiums.append(0)
            taxes.append(0)
    
    dataframe['Total Pro-Rata GWP'] = pro_rata_gwps
    dataframe['Total Earned Premium'] = earned_premiums
    dataframe['Total Unearned Premium'] = unearned_premiums
    dataframe['Total Taxes'] = taxes


# Reads the file, adds calculated columns to dataframe, aggregates the data by Company Name, and writes to new file.

def main():
    report_date = '2022-08-01'
    df = read_file()
    # your processing here......
    add_calculated_columns(df, report_date)
    # aggregates
    aggr = df.groupby('Company Name')[["Total Pro-Rata GWP", "Total Earned Premium", "Total Unearned Premium", "Total Taxes"]].sum()
    print("Aggregates")
    print(aggr)
    
    write_new_file(df, report_date)
    return

if __name__ == "__main__":
    main()