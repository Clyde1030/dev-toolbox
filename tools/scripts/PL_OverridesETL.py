## Script to test loading in overrides data using Python/Pandas and BCP ##
## Code below can be copy-pasted and edited to work with other override files ##

import pandas as pd
import re

def main():
    ## Filename Input and Processing ##
    filename = "O-1111-LIF-2022-12-123V.txt"
    pattern = "(?P<type>(O|C))-(?P<carrierId>\d{4})-(?P<lob>\w*)-(?P<year>\d{4})-(?P<month>\d{2})-(.*)\.txt$"
    match = re.match(pattern, filename)

    ## Headers to match PL input file ##
    headers = [
        'PolicyNumber',
        'TransactionDate',
        'IssueDate',
        'CarrierProductID',
        'DurationText',
        'PremiumText',
        'Prem',
        'CommissionRateA',
        'OverrideRate',
        'BucketNumber',
        'SplitPercentage',
        'CommissionAdjustmentFactor',
        'CommissionA',
        'Override',
        'InsuredName',
        'Commission_Original',
        'PolicyStatus',
        'CarrierProducerID',
        'CarrierProducerName',
        'CarrierContracteeID',
        'CarrierContracteeName',
        'PaymentMode',
        'DurationOLD',
        'PaymentModeType'
    ]

    ## Read in the input file ##
    df = pd.read_csv(filename, delimiter=';', header=None, names=headers)

    ## Code for applying logic to columns ##
    df["Fixed/Var"] = df.apply(lambda x: "MFSM" if x["PolicyNumber"][0:2] in ("VP","VM","SV") else "MFH", axis=1)
    df["SourceFileName"] = filename.split('.')[0] + df["Fixed/Var"]
    df["CarrierID"] = match.group('carrierId')
    df["MemFirmID"] = 0
    df["ProductID"] = 0
    df["YearApplied"] = match.group('year')
    df["MonthApplied"] = match.group('month')
    df["ProducerID"] = 0
    df["RiskNumber"] = ""
    df["RiskName"] = ""
    df["CarrierProductName"] = df['CarrierProductID']
    df["ExcessPremium"] = df.apply(lambda x: x["Prem"] if x["PremiumText"]!="Trail" and x["OverrideRate"]<=.01 else 0, axis=1)
    df["TrailPremium"] = df.apply(lambda x: x["Prem"] if x["PremiumText"]=="Trail" else 0, axis = 1)
    df["PaidPremium"] = df.apply(lambda x: x["Prem"] if x["ExcessPremium"]==0 and x["TrailPremium"]==0 else 0, axis=1)
    df["Commission"] = df.apply(lambda x: x["Commission_Original"] if match.group("type") == "C" else 0, axis=1)
    df["MandM"] = 0
    df["SystemFee"] = 0
    df["PercentAppliedID"] = ""
    df["CarrierBrokerID"] = df["CarrierContracteeID"].str.replace(' ', '')
    df["CarrierBrokerName"] = df["CarrierContracteeName"].str.replace(' ', '')
    df["CommissionOption"] = df.apply(lambda x: x["CarrierProductID"][-1] , axis = 1).str.replace(' ', '')
    df["TransactionDate"] = pd.to_datetime(df["TransactionDate"])
    df["IssueDate"] = pd.to_datetime(df["IssueDate"])
    df["Duration"] = round(((df["TransactionDate"] - df["IssueDate"]).dt.days)/365.25 + 1, 0)
    df["Duration"] = df["Duration"].astype(int)
    df["TransactionDate"] = pd.to_datetime(df["TransactionDate"]).dt.strftime("%Y-%m-%d")
    df["IssueDate"] = pd.to_datetime(df["IssueDate"]).dt.strftime("%Y-%m-%d")
    df["CommissionRate"] = df["CommissionRateA"]
    df["InsuredName"] = df["InsuredName"].str.replace(',', '')
    df["DurationText"] = df["DurationText"].str.replace(' ', '')
    df["PremiumText"] = df["PremiumText"].str.replace(' ', '')
    df["PremiumText2"] = df.apply(lambda x: "TRL" if df["PremiumText"].str.split(' ')[0] == "Trail" else df["PremiumText"].str.split(' ')[0], axis=1)
    df["RevenueCategory"] = df.apply(
        lambda x: "OVERFY" if x["PremiumText2"] == "FY" and x["BucketNumber"] == 1 else (
            "OVERFY-X" if x["PremiumText2"] == "FY" and x["BucketNumber"] > 1 else (
                "OVERRN" if x["PremiumText2"] == "RL" and x["BucketNumber"] == x["Duration"] else (
                    "OVERRN-X" if x["PremiumText2"] == "RL" and x["BucketNumber"] > x["Duration"] else "OVERTR"
                )
            )
        ),
        axis=1
    )

    ## Columns that need to be exported to load with BCP ##
    ## Order matters! ##
    exports_cols = [
        'SourceFileName',
        'CarrierID',
        'MemFirmID',
        'CarrierContracteeID',
        'CarrierContracteeName',
        'ProductID',
        'CarrierProductID',
        'YearApplied',
        'MonthApplied',
        'ProducerID',
        'CarrierProducerID',
        'CarrierProducerName',
        'PolicyNumber',
        'IssueDate',
        'InsuredName',
        'RiskNumber',
        'RiskName',
        'TransactionDate',
        'Duration',
        'PaidPremium',
        'ExcessPremium',
        'TrailPremium',
        'Commission',
        'Override',
        'MandM',
        'SystemFee',
        'OverrideRate',
        'PaymentMode',
        'PaymentModeType',
        'DurationText',
        'PercentAppliedID',
        'SplitPercentage',
        'CarrierProductName',
        'CarrierBrokerID',
        'CarrierBrokerName',
        'PolicyStatus',
        'CommissionRate',
        'CommissionAdjustmentFactor',
        'RevenueCategory',
        'BucketNumber',
        'CommissionOption'
    ]

    ## Exports file ##
    export_name = filename.split('.txt')[0] + '.csv'
    df.to_csv(export_name, columns=exports_cols, index=False)

    return

if __name__ == "__main__":
    main()