Namespace BPS_BL.BPS

    Public Enum InteractionSource
        '??? Not sure about these
        Unknown = 0
        System = 100
        Phone = 1
        Email = 2
        PostalMail = 3
        Online = 4
        Fax = 5
    End Enum

    Public Enum AnswerType
        Rating = 1
        YesNo = 2
        Response = 3
        YesNoWithResponse = 4
    End Enum


    Public Enum ResourceType
        Company = 1
        Brand = 2
        Series = 3
        Product = 4
        General = 5
        Part = 6
        Category = 7
        CustomerCase = 8
    End Enum

    Public Enum ImageSize
        Feature = 0
        Small = 1
        Medium = 2
        Large = 3
    End Enum
    Public Enum ImageType
        Alternate = 1
        DataCube = 2

    End Enum

    Public Enum CustomerAccountAuthenticationMode
        CustomerAccountID = 1
        UsernameAndPassword = 2
        CustomerAndAccountID = 3
        UsernameAndAccountID = 4
    End Enum
    Public Enum KnowledgeType
        All = 0
        ProductGuides = 1
        MaintenanceAndCare = 2
        Troubleshooting = 3
        RecallsAndAlerts = 4
        NewsAndInformation = 5
        Search = 6 'Internally used for search tab
        CookingGuides = 7
        PartsAndAccessories = 8 'Internally used for Parts & Accessories
        Specifications = 9 'Internally used for Specs

    End Enum
    Public Enum ManageKnowledgeMode
        'These are used internally - this info is not in the database
        '..it is used to store information on the available sections of 
        'knowledge available to manage - Depending upon your mode
        'you may have different values available to modify.
        'These do not always directly relate to the above Knowledge Type 
        '.. for example you may add Product Knowledge (1) and it's knowledge type 
        'can be any of the above knowledge types

        ProductKnowledge = 1
        MaintenanceAndCare = 2
        RecallsAndAlerts = 3
        CompanyNews = 4
        CookingGuides = 5
        Troubleshooting = 6
        ProductGuides = 7
    End Enum

    Public Enum BPSItemType
        Part = 0
        Product = 1
    End Enum

    Public Enum EmailPriorities
        Normal = 0
        Low = 1
        High = 2
    End Enum


    Public Enum MailPriority
        Normal = 0
        Low = 1
        High = 2
    End Enum

    Public Enum EmailBodyFormat
        Text = 0
        HTML = 1

    End Enum

    Public Enum ProductSeriesAttributeType
        Bullet = 1
        DropShip = 2
        TruckDelivery = 3
        CompanyExclusive = 4
        Oversized = 5
        ActiveItemMktgWebsite = 101
        CurrentYearProduct = 102

    End Enum

    Public Enum ProductAttributeType
        Bullet = 1
        DropShip = 2
        TruckDelivery = 3
        CompanyExclusive = 4
        Oversized = 5
        Length = 6
        Width = 7
        Height = 8
        QtyPer = 9
        Scent = 10
        Weight = 11
        ActiveItemMktgWebsite = 101
        CurrentYearProduct = 102

    End Enum
    Public Enum ProductSeriesAssociatedType
        CrossSell = 1
        UpSell = 2
        Bundled = 3

    End Enum

    Public Enum PurchaseStatus
        NoInvoice_Web = 1
        Invoice_Web = 2
        NoWeb_Invoice = 3
        NoInvoiceNoWeb = 4

    End Enum
    Public Enum CreditApplicationStatus
        Unknown = 0
        Approved = 1
        Pending = 2
        Declined = 3
    End Enum
    Public Enum AccountStatus
        Unknown = 0
        Approved = 1
        Pending = 2
        Declined = 3
        ApprovedwithPendingChanges = 4
        TBD5 = 5
        TBD6 = 6

    End Enum

    Public Enum EmailTemplate
        BlankTemplate = 0
        NewSurvey = 1
        PasswordReminder = 2
        CaseUpdate = 3
        POPRequest = 4
        POPThankyou = 5
        POPDecline = 6
        ProductRegistration = 8
        CustomEmail = 9
        PasswordReset = 10
        ContactUsWarrantyOrder = 11
        POPReceivedNotValid = 12
        POPUnclear = 13
        ReferraltoConsumerRelations = 14
        RequestforPOPEmail = 15
        POPNotValid = 16
        NewCase = 17
        EmployeeRouteCase = 18
        NewCustomerAccountConfirmation = 19
        CaseClosed = 20
        NonClaimPOPRequest = 21
        PasswordResetCheckout = 22
        RegistrationValidationEmail = 23
        RegistrationStatusEmail = 24
        RegistrationConfirmation = 25




    End Enum







#Region "DataCube Enum_Section"

    Public Enum DataItemType

        Category = 1
        VirtualCategory = 2
        Product = 3
        Image = 6
        Hyperlink = 7
        Text = 8
        Image_Flash = 9
        promotion = 10

    End Enum

    Public Enum DataCubeType
        'Category = 1
        'VirtualCategory = 2
        'Product = 3
        'Text = 4
        'Image_Flash = 6
        'Hyperlink = 7
        Category = 1
        VirtualCategory = 2
        Product = 3
        Image = 6
        Hyperlink = 7
        Text = 8
        Image_Flash = 9
        promotion = 10

    End Enum

#End Region

End Namespace
