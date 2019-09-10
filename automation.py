import argparse
import pandas as pd
import numpy as np

def automation(input_filename, output_filename, show_print_lines=False):
    
    # load excel sheets into data frame
    # 'data/Compiled Index.xlsx'
    if show_print_lines:
        print("Load excel sheet named {0} into dataframe..".format(input_filename))
    dfs = pd.read_excel(input_filename, sheet_name=None)

    # removing none values with 0
    for df in dfs:
        dfs[df] = dfs[df].replace(0,np.NaN)
        dfs[df] = dfs[df].dropna(how='all')
        dfs[df] = dfs[df].fillna(0)

    # Consistant shape and size for better calculation
    dfs['ShareNumber']= dfs['ShareNumber'][dfs['ShareNumber'].columns[:-1]]
    dfs['CDividend']= dfs['CDividend'][dfs['CDividend'].columns[:-1]]
    dfs['Multiplier']= dfs['Multiplier'][dfs['Multiplier'].columns[:-1]]
    dfs['RightP']= dfs['RightP'][dfs['RightP'].columns[:-1]]
    dfs['Price'] = dfs['Price'][dfs['ShareNumber'].columns]

    # calculation m cap using the formula
    if show_print_lines:
        print("Calculating m_cap..")
    mcap_df = dfs['Price'][dfs['Price'].columns[2:]] * dfs['ShareNumber'][dfs['ShareNumber'].columns[2:]]
    # adding sector and ticker column
    mcap_df.insert(loc=0, column='Sector', value=dfs['Price']['Sector'].values)
    mcap_df.insert(loc=0, column='Ticker', value=dfs['Price']['Ticker'].values)

    # sector_wise mcap value
    sector_wise_market_cap = mcap_df.groupby("Sector")[mcap_df.columns].sum()

    # calculate adj_div and adj_right using formulas
    if show_print_lines:
        print("calculating adj_div and adj_right using formulas")
    adj_div_df = dfs['CDividend'][dfs['CDividend'].columns[2:]] * dfs['ShareNumber'][dfs['ShareNumber'].columns[2:]]
    adj_right_df = dfs['Multiplier'][dfs['Multiplier'].columns[2:]] * dfs['RightP'][dfs['RightP'].columns[2:]] * dfs['ShareNumber'][dfs['ShareNumber'].columns[2:]]

    adj_div_df.insert(loc=0, column='Sector', value=dfs['Price']['Sector'].values)
    adj_div_df.insert(loc=0, column='Ticker', value=dfs['Price']['Ticker'].values)
    adj_right_df.insert(loc=0, column='Sector', value=dfs['Price']['Sector'].values)
    adj_right_df.insert(loc=0, column='Ticker', value=dfs['Price']['Ticker'].values)

    # Calculating TDivisor using formula, here the portion of (mcapAW - addvAW + adjrAW)/ mcapAW
    if show_print_lines:
        print("Calculating TDivisor using formula, here the portion of (mcapAW - addvAW + adjrAW)/ mcapAW..")
    t_divisor_2nd_part = mcap_df[mcap_df.columns[2:]] - adj_div_df[adj_div_df.columns[2:]] + adj_right_df[adj_right_df.columns[2:]]
    t_divisor_2nd_part = t_divisor_2nd_part / mcap_df[mcap_df.columns[2:]]

    # fill null values with 0
    t_divisor_2nd_part = t_divisor_2nd_part.fillna(0)

    # now the whole part of the formula, =AV*((mcapAW - addvAW + adjrAW)/ mcapAW)
    if show_print_lines:
        print("now calculating the whole part of the formula, =AV*((mcapAW - addvAW + adjrAW)/ mcapAW")
    t_divisor = pd.DataFrame(dfs['TDivisor'][dfs['TDivisor'].columns[2:-1]].values * t_divisor_2nd_part[t_divisor_2nd_part.columns[1:]].values, columns=dfs['TDivisor'].columns[2:-1], index=dfs['TDivisor'].index)
    t_divisor.insert(loc=0, column='Sector', value=dfs['Price']['Sector'].values)
    t_divisor.insert(loc=0, column='Ticker', value=dfs['Price']['Ticker'].values)

    # t_return calculation
    if show_print_lines:
        print("t_return calculation")
    t_return = mcap_df[mcap_df.columns[3:]] / t_divisor[t_divisor.columns[2:]]

    # finding infinity indexes
    inf_idx = t_return.index[np.isinf(t_return).any(1)]

    # insert missing date column
    t_return.insert(loc=0, column=mcap_df.columns[2], value=dfs['TReturn'][dfs['TReturn'].columns[2]].values)

    # replace with first column value for inf
    for index in inf_idx:
        t_return.loc[index,t_return.columns] = t_return.at[index, t_return.columns[0]]

    t_return = t_return.fillna(0)

    sector_wise_cash_divided = adj_div_df.groupby("Sector")[mcap_df.columns].sum()
    sector_wise_right_share = adj_right_df.groupby("Sector")[mcap_df.columns].sum()

    # prepare the output for given sectors
    if show_print_lines:
        print("prepare the output for given sectors")
    list_df = []
    for name in ['Bank', 'NBFI', 'Pharmaceuticals']:
        
        IndexData = {
            name + ' Market Cap' : sector_wise_market_cap.loc[name].values,
            'Cash Dividend' : sector_wise_cash_divided.loc[name].values,
            'Right Share' : sector_wise_right_share.loc[name].values,
            
        }
        list_df.append(IndexData)

    # these could be improved, No time to restructure
    # bank
    bank_df = pd.DataFrame(list_df[0], index=mcap_df.columns[2:])
    bank_df['Divisor'] = 283
    bank_df['Adjusments'] = 0
    bank_df['Bank'] = bank_df['Bank Market Cap'].values/bank_df['Divisor'].values
    bank_df = bank_df.transpose()

    nbfi_df = pd.DataFrame(list_df[1], index=mcap_df.columns[2:])
    nbfi_df['Divisor'] = 72
    nbfi_df['Adjusments'] = 0
    nbfi_df['NBFI'] = nbfi_df['NBFI Market Cap'].values/nbfi_df['Divisor'].values
    nbfi_df = nbfi_df.transpose()

    Pharmaceuticals_df = pd.DataFrame(list_df[2], index=mcap_df.columns[2:])
    Pharmaceuticals_df['Divisor'] = 130
    Pharmaceuticals_df['Adjusments'] = 0
    Pharmaceuticals_df['Pharmaceuticals'] = Pharmaceuticals_df['Pharmaceuticals Market Cap'].values/Pharmaceuticals_df['Divisor'].values
    Pharmaceuticals_df = Pharmaceuticals_df.transpose()

    final_output = bank_df.append(nbfi_df).append(Pharmaceuticals_df)
    # export to csv into folder data
    # 'data/final_output.csv'
    if show_print_lines:
        print("export dataframe to a output csv file named {0}..".format(output_filename))
    export_csv = final_output.to_csv (output_filename)



if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("inputfile", type=str,
                        help="input excel file name")
    parser.add_argument("outfile", type=str,
                        help="output csv file name")

    parser.add_argument("-v", "--verbose", action="store_true",
                        help="increase output verbosity")

    args = parser.parse_args()

    
    
    if args.verbose:
        print("yes")
        automation(input_filename=args.inputfile, output_filename=args.outfile, show_print_lines=True)
    else:
        print("no  verbosity")
        automation(input_filename=args.inputfile, output_filename=args.outfile)
