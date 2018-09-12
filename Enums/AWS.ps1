Add-Type -TypeDefinition @"
    public enum AWS {
        AWSEC2Details,
        AWSRDSDetails,
        AWSLBDetails,
        AWSNetworkDetails,
        AWSEIPDetails,
        AWSIAMDetails
    }
"@