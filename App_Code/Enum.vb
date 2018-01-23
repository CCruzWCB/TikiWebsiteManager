Namespace Management

    Public Enum EmailPriorities
        Normal = 0
        Low = 1
        High = 2
    End Enum

    Public Enum HeaderDisplayMode
        User = 1
        UserAdmin = 2
        VendorAdmin = 3
        RetailerAdmin = 4
        ApplicationAdministrator = 5
    End Enum

    Public Enum CustomField
        ReasonCode = 15
        ProblemModel = 16
        WarrantyItemNoCost = 17
        AppeasementOrder = 18
        PONumber = 23
        StoreNumber = 24
        AddressType = 19
        TaxExemptNumber = 20

    End Enum
    Public Enum SubMenu

        NoMenu = 0

        RetailerSubMenu1 = 1
        RetailerSubMenu2 = 2
        RetailerSubMenu3 = 3

        VendorSubMenu1 = 4
        VendorSubMenu2 = 5
        VendorSubMenu3 = 6

        UserSubMenu1 = 7
        UserSubMenu2 = 8
        UserSubMenu3 = 9

        AdminSubMenu1 = 10
        AdminSubMenu2 = 11
        AdminSubMenu3 = 12


    End Enum
    Public Enum HeaderMenuOptions
        User = 1
        UserAdmin = 2
        VendorAdmin = 3
        RetailerAdmin = 4
        ApplicationAdministrator = 5
    End Enum
    Public Enum ApplicationRoles
        ApplicationUser = 1
        ApplicationUserManager = 2
        ApplicationVendorManager = 3
        ApplicationRetailerManager = 4
        ApplicationAdministrator = 5

    End Enum


    Public Enum ShoppingCartStatus
        NewBasketNoItems = 0
        NewBasketWithItems = 1
        NewBasketStartCheckout = 2
        NewBasketCheckoutStart = 3
        NewBasketCheckoutEnd = 4
        NewBasketPaymentPending = 5
        NewBasketPaymentComplete = 6
    End Enum

    Public Enum MessageType
        GeneralMessage = 1
        SyntaxMessage = 2
        ErrorMessage = 3
        WizardMessage = 4
        instructionMessage = 5
        ConfirmationMessage = 6
    End Enum


    Public Enum OrderStatus
        'Submitted = 0
        'Approved = 1
        'Denied = 2
        'SenttoOrderProcessingSystem = 3
        'ReceivedfromOrderProcessingSystem = 4
        'OrderInfoNotReceivedfromOrderProcessingSystem = 5
        'ErrorSendingOrder = 6
        'CancelbyUser = 8
        'Complete = 7
        Archived = -1
        Pending = 0
        Clearing = 2
        BackedOrderedAlternative = 3
        Review = 4
        OnHold = 5
        Inprocess = 6
        Waiting = 7
        Released = 10
        BackOrdered = 20
        WareHouse = 30
        Shipped = 40
        Returned = 50
        PendingReturn = 52
        Refunded = 60
        Cancelled = 90
        AutoCancelled = 95

    End Enum

    Public Enum OrderBatchStatus
        Archived = -1
        Pending = 0
        Clearing = 2
        BackedOrderedAlternative = 3
        Review = 4
        OnHold = 5
        Inprocess = 6
        Waiting = 7
        Released = 10
        BackOrdered = 20
        WareHouse = 30
        Shipped = 40
        Returned = 50
        PendingReturn = 52
        Refunded = 60
        Cancelled = 90
        AutoCancelled = 95
    End Enum

    Public Enum EmailBodyFormat
        Text = 0
        HTML = 1

    End Enum

    Public Enum ElementType
        Image = 1
        Document = 2
        ZipFile = 3
        UnknownType = 4

    End Enum

    Public Enum NoModelReasonCode
        ModelNotInSystem = 1
        CustomerUnableToLocate = 2
        CustomerNoAccessToGrill = 3
        NotNeededGeneralInfo = 4
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
    Public Enum RegistrationStatus
        NotRegistered = 0
        Registered = 1
        Pending = 2
        Authorized = 3
        Denied = 4
        Ignored = 5
    End Enum

    Public Enum OrderSearchType
        Today = 0
        Yesterday = 1
        Weekly = 2
        YTD = 3
        Month = 9
        OrderName = 4
        DateRange = 5
        OrderID = 6
        Store = 7
        AllStoreOrders = 8


    End Enum
    Public Enum RegistrationAuthenticationType
        InvalidDomain = -1 'All email address in domain can register 
        PublicRegistration = 1 'All email address in domain can register 
        PrivatePreRegistrationRequired = 2 'Email address or basic information must be setup by the administrator
        VerificationRequired = 3 'All email address can register but must be approval by administrator
    End Enum

    Public Enum ReasonCodeType
        Denied = 0
        Approved = 1
        Ignored = 2
    End Enum
    Public Enum RatingScale
        Poor_to_Great = 0
        Never_to_Always = 1
        PartList = 2
    End Enum
End Namespace
