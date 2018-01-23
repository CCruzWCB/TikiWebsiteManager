
/*************************************************************************************************************************
	Updated to include a Sequence field in the Product Table
 *************************************************************************************************************************/
USE [LLFBPS]
GO

/****** Object:  Table [dbo].[Product]    Script Date: 03/04/2011 14:44:50 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO
IF NOT EXISTS(SELECT * FROM sys.columns WHERE Name = N'Sequence' AND Object_ID = Object_ID(N'[dbo].[Product]'))
BEGIN 
	ALTER TABLE [dbo].[Product] ADD [Sequence] [int] 
	CONSTRAINT [DF_Product_Sequence]  DEFAULT ((0)) NOT NULL 
END
GO

/*************************************************************************************************************************
	Updated to include the Sequence Field in the vProducts View
 *************************************************************************************************************************/
USE [LLFBPS]
GO

/****** Object:  View [dbo].[vProducts]    Script Date: 03/04/2011 14:58:41 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

ALTER VIEW [dbo].[vProducts]
AS
SELECT     dbo.Product.ProductID
		, (CASE WHEN CharIndex('.', dbo. Product .ProductModelNumber) > 0 
			AND dbo. Product .IsSupported = 1 THEN 
				LEFT(dbo.Product.ProductModelNumber, Charindex('.', dbo.Product.ProductModelNumber) - 1) 
            ELSE dbo. Product .ProductModelNumber END) AS productmodelnumber
        , dbo.Product.Name
        , dbo.Product.Description
        , dbo.Product.Year
        , dbo.Product.Freight
        , ISNULL(dbo.Product.MSRP, 0) AS MSRP
        , dbo.Product.Specifications
        , dbo.Product.PrimaryResourceImageID
        , dbo.Product.PrimaryResourceImageID_large
        , dbo.Product.PrimaryResourceImageID_small
        , dbo.Product.PrimaryResourceImageID_feature
        , dbo.Image.ImagePath
        , Image1.ImagePath AS ImagePath_small
        , Image2.ImagePath AS ImagePath_large
        , Image3.ImagePath AS ImagePath_feature
        , dbo.Product.CreatedName
        , dbo.Product.DateCreated
        , dbo.Product.IsSupported
        , dbo.Product.CreatedBy
        , dbo.Product.IsSellable
        , dbo.Product.ProductSeriesID
        , dbo.Product.Active
        , dbo.ProductCategoryJUNC.ParentCategoryID
        , dbo.ProductSeries.ProductCategoryID
        , dbo.ProductSeries.BrandID
        , dbo.Brand.Name AS BrandName
        , dbo.ProductSeries.Name AS SeriesName
        , ProductCategory2.Name AS ParentCategoryName
        , dbo.ProductCategory.Name AS CategoryName
        , dbo.ProductSeries.ECommerceCategoryID
        , ProductCategory_2.ProductCategoryID AS EcommerceParentCategoryID
        , ProductCategory_1.Name AS ECommerceCategoryName
        , ProductCategory_2.Name AS EcommerceParentCategoryName
        , dbo.Product.AdditionalShipping
        , dbo.Product.AvailableQuantity
        , dbo.Product.ExpectedDate
        , dbo.ProductDiagram.FilePath + dbo.ProductDiagram.FileName AS DiagramFilePath
        , dbo.Product.Sequence
FROM	dbo.Image AS Image1 WITH (NOLOCK) 
		RIGHT OUTER JOIN dbo.ProductDiagram WITH (NOLOCK) 
		RIGHT OUTER JOIN dbo.Product WITH (NOLOCK) 
		INNER JOIN dbo.ProductSeries WITH (NOLOCK) 
			ON dbo.Product.ProductSeriesID = dbo.ProductSeries.ProductSeriesID 
		INNER JOIN dbo.Brand WITH (NOLOCK) 
			ON dbo.ProductSeries.BrandID = dbo.Brand.BrandID 
			ON dbo.ProductDiagram.ProductID = dbo.Product.ProductID 
		LEFT OUTER JOIN dbo.ProductCategory WITH (NOLOCK) 
			ON dbo.ProductSeries.ProductCategoryID = dbo.ProductCategory.ProductCategoryID 
		LEFT OUTER JOIN dbo.ProductCategoryJUNC WITH (NOLOCK) 
			ON dbo.ProductSeries.ProductCategoryID = dbo.ProductCategoryJUNC.ProductCategoryID 
		LEFT OUTER JOIN dbo.ProductCategory AS ProductCategory2 WITH (NOLOCK) 
			ON dbo.ProductSeries.ProductCategoryID = dbo.ProductCategoryJUNC.ProductCategoryID 
			AND ProductCategory2.ProductCategoryID = dbo.ProductCategoryJUNC.ParentCategoryID 
		LEFT OUTER JOIN dbo.ProductCategory AS ProductCategory_1 WITH (NOLOCK) 
			ON dbo.ProductSeries.ECommerceCategoryID = ProductCategory_1.ProductCategoryID 
		LEFT OUTER JOIN dbo.ProductCategoryJUNC AS ProductCategoryJUNC_1 WITH (NOLOCK) 
		LEFT OUTER JOIN dbo.ProductCategory AS ProductCategory_2 WITH (NOLOCK) 
			ON ProductCategoryJUNC_1.ParentCategoryID = ProductCategory_2.ProductCategoryID 
			ON dbo.ProductSeries.ECommerceCategoryID = ProductCategoryJUNC_1.ProductCategoryID 
		LEFT OUTER JOIN dbo.Image WITH (NOLOCK) 
			ON dbo.Product.PrimaryResourceImageID = dbo.Image.ImageID 
		LEFT OUTER JOIN dbo.Image AS Image2 WITH (NOLOCK) 
			ON dbo.Product.PrimaryResourceImageID_large = Image2.ImageID 
			ON Image1.ImageID = dbo.Product.PrimaryResourceImageID_small 
		LEFT OUTER JOIN dbo.Image AS Image3 WITH (NOLOCK) 
			ON dbo.Product.PrimaryResourceImageID_feature = Image3.ImageID

GO

/*************************************************************************************************************************
	Updated to include the Sequence Field in the vProductCategories View
 *************************************************************************************************************************/
USE [LLFBPS]
GO

/****** Object:  View [dbo].[vProductCategories]    Script Date: 03/04/2011 15:01:12 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

ALTER VIEW [dbo].[vProductCategories]
AS
SELECT	dbo.ProductCategory.ProductCategoryID
		, dbo.ProductCategory.Name
		, dbo.ProductCategory.Description
		, dbo.ProductCategory.PrimaryResourceImageID
		, dbo.ProductCategory.CreatedBy
		, E.FirstName + ' ' + E.LastName AS CreatedName
		, dbo.ProductCategory.DateCreated
		, dbo.ProductCategory.Active
		, dbo.ProductCategoryJUNC.ParentCategoryID
		, dbo.Image.ImagePath
		, ProductCategory_1.Name AS ParentCategoryName
		, dbo.ProductCategory.IsEcommerceCategory
		, dbo.ProductCategory.BrandID
		, dbo.Brand.Name AS BrandName
		, dbo.Brand.Name + '-' + dbo.ProductCategory.Name AS DisplayCategoryName
		, ISNULL(CategorySequence.Sequence, 0) AS Sequence
		, ISNULL(CategorySequence_1.Sequence, 0) AS ParentSequence
FROM	dbo.ProductCategory 
		INNER JOIN dbo.ProductCategoryJUNC 
			ON dbo.ProductCategory.ProductCategoryID = dbo.ProductCategoryJUNC.ProductCategoryID 
		INNER JOIN dbo.ProductCategory AS ProductCategory_1 
			ON dbo.ProductCategoryJUNC.ParentCategoryID = ProductCategory_1.ProductCategoryID 
		INNER JOIN dbo.Brand 
			ON dbo.ProductCategory.BrandID = dbo.Brand.BrandID 
		LEFT OUTER JOIN LLFWebSite.dbo.CategorySequence AS CategorySequence_1 
			ON ProductCategory_1.ExternalCategoryID = CategorySequence_1.ExternalCategoryID 
		LEFT OUTER JOIN LLFWebSite.dbo.CategorySequence 
			ON dbo.ProductCategory.ExternalCategoryID = CategorySequence.ExternalCategoryID 
		LEFT OUTER JOIN dbo.Image 
			ON dbo.ProductCategory.PrimaryResourceImageID = dbo.Image.ImageID 
		LEFT OUTER JOIN LLFCMS.dbo.Employee AS E 
			ON dbo.ProductCategory.CreatedBy = E.EmployeeID

GO

/*************************************************************************************************************************
	Updated GetEcommerceMenuByBrand to include Sequence information
 *************************************************************************************************************************/

USE [LLFBPS]
GO

/****** Object:  StoredProcedure [dbo].[GetEcommerceMenuByBrand]    Script Date: 03/04/2011 15:32:03 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[GetEcommerceMenuByBrand]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[GetEcommerceMenuByBrand]
GO

USE [LLFBPS]
GO

/****** Object:  StoredProcedure [dbo].[GetEcommerceMenuByBrand]    Script Date: 03/04/2011 15:32:03 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		Corey Perrymond
-- Create date: 05/06/08
-- Description:	Get Ecommerce Category Menu By Brand
-- =============================================
CREATE PROCEDURE [dbo].[GetEcommerceMenuByBrand]-- 2  -- 1 
	-- Add the parameters for the stored procedure here
	@BrandID INT = 1 

AS

BEGIN
	SET NOCOUNT ON;

DECLARE @SubCategories TABLE 
(ProductCategoryID INT, 
[Name] varchar(50), 
[Description] varchar(200),
ParentCategoryID INT,
Roles INT,
Sequence INT, 
URL varchar(500), 
thumbnail varchar(200) 
)

DECLARE @Products TABLE 
(ProductCategoryID INT, 
[Name] varchar(50), 
[Description] varchar(200),
ParentCategoryID INT,
Roles INT,
Sequence INT, 
URL varchar(500), 
thumbnail varchar(200) 
)

INSERT INTO @SubCategories 
EXECUTE dbo.GetSequencedSubCategories @BrandID

INSERT INTO @Products
EXECUTE dbo.GetProductSequence @BrandID

/* Root*/
SELECT 
	LLFBPS.dbo.ProductCategory.ProductCategoryID, 
	'Home' AS Name, 
	'Home' AS Description, 
	LLFBPS.DBO.ConvertNoCategory(LLFBPS.dbo.ProductCategoryJUNC.ParentCategoryID,LLFBPS.dbo.ProductCategory.ProductCategoryID) as ParentCategoryID,
---1 as ParentCategoryID,
	NULL AS Roles, 
	0 AS Sequence, 
	'~/Consumer/default.aspx?CategoryID=' + convert(varchar(5),LLFBPS.dbo.ProductCategory.ProductCategoryID) AS url , 
	'' AS thumbnail 
FROM	LLFBPS.dbo.ProductCategory 
	INNER JOIN LLFBPS.dbo.ProductCategoryJUNC 
		ON LLFBPS.dbo.ProductCategory.ProductCategoryID = LLFBPS.dbo.ProductCategoryJUNC.ProductCategoryID 
WHERE	LLFBPS.dbo.ProductCategory.ProductCategoryID = 0

UNION 


--/* Brands*/
--SELECT 
--	BrandID * 1000 as ProductCategoryID, 
--	Name, 
--	isnull(Description,'no description') as Description, 
--	(0) as ParentCategoryID,
--
--	NULL AS Roles, 
--1 as Sequence, 
--	'~/CustomerHome/BrandListings.aspx?BrandID=' + convert(varchar(5),BrandID) as url 
--FROM 
--	Brand Where active = 1

--UNION 

/* Top Level Ecommerce Categories */
/* Add Tiki Link */
SELECT	LLFBPS.dbo.ProductCategory.ProductCategoryID
	, 'Tiki® Products' AS [Name]
	, 'Tiki® Products' AS Description
	, (0) AS ParentCategoryID
	, NULL AS Roles
	, 1 AS Sequence
	, '~/Consumer/TikiProducts.aspx' AS url 
	, '' AS thumbnail 
FROM	LLFBPS.dbo.ProductCategory 
	INNER JOIN LLFBPS.dbo.ProductCategoryJUNC 
		ON LLFBPS.dbo.ProductCategory.ProductCategoryID = LLFBPS.dbo.ProductCategoryJUNC.ProductCategoryID 
WHERE LLFBPS.dbo.ProductCategoryJUNC.ParentCategoryID = 0 
	AND LLFBPS.dbo.ProductCategory.productcategoryid > 0 
	AND LLFBPS.dbo.ProductCategory.active = 1
	AND LLFBPS.dbo.ProductCategory.BrandID = @BrandID 
	AND LLFBPS.dbo.ProductCategory.IsEcommerceCategory = 1 
	AND LLFBPS.dbo.ProductCategory.[name] LIKE 'Tiki%'

UNION 
/* Add Lamplight Link */
SELECT	LLFBPS.dbo.ProductCategory.ProductCategoryID
	, 'Lamplight® Products' AS [Name]
	, 'Lamplight® Products'  AS Description
	, (0) AS ParentCategoryID
	, NULL AS Roles
	, 2 AS Sequence
	, '~/Consumer/LamplightProducts.aspx' AS url 
	, '' AS thumbnail 
FROM	LLFBPS.dbo.ProductCategory 
	INNER JOIN LLFBPS.dbo.ProductCategoryJUNC 
		ON LLFBPS.dbo.ProductCategory.ProductCategoryID = LLFBPS.dbo.ProductCategoryJUNC.ProductCategoryID 
WHERE LLFBPS.dbo.ProductCategoryJUNC.ParentCategoryID = 0 
	AND LLFBPS.dbo.ProductCategory.productcategoryid > 0 
	AND LLFBPS.dbo.ProductCategory.active = 1
	AND LLFBPS.dbo.ProductCategory.BrandID = @BrandID 
	AND LLFBPS.dbo.ProductCategory.IsEcommerceCategory = 1 
	AND LLFBPS.dbo.ProductCategory.[name] LIKE 'Lamplight%'

UNION 
/* Add AromoGlow Link */
SELECT	LLFBPS.dbo.ProductCategory.ProductCategoryID
	, 'AromoGlow® Products' AS [Name]
	, 'AromoGlow® Products'  AS Description
	, (0) AS ParentCategoryID
	, NULL AS Roles
	, 3 AS Sequence
	, '~/Consumer/AromaGlowProducts.aspx' AS url
	, '' AS thumbnail 
FROM	LLFBPS.dbo.ProductCategory 
	INNER JOIN LLFBPS.dbo.ProductCategoryJUNC 
		ON LLFBPS.dbo.ProductCategory.ProductCategoryID = LLFBPS.dbo.ProductCategoryJUNC.ProductCategoryID 
WHERE LLFBPS.dbo.ProductCategoryJUNC.ParentCategoryID = 0 
	AND LLFBPS.dbo.ProductCategory.productcategoryid > 0 
	AND LLFBPS.dbo.ProductCategory.active = 1
	AND LLFBPS.dbo.ProductCategory.BrandID = @BrandID 
	AND LLFBPS.dbo.ProductCategory.IsEcommerceCategory = 1 
	AND LLFBPS.dbo.ProductCategory.[name] LIKE 'AromaGlow%'

UNION 
/* Ecommerce Sub Categories */
SELECT	ProductCategoryID
		, [Name]
		, [Description]
		, ParentCategoryID
		, Roles
		, [Sequence]
		, URL
		, thumbnail 
FROM @SubCategories

UNION 
SELECT 
	-3 AS ProductCategoryID
	, 'How to Use' AS [Name]
	, 'How to Use' AS Description
	, ISNULL((SELECT LLFBPS.dbo.ProductCategory.ProductCategoryID 
		FROM LLFBPS.dbo.ProductCategory WITH (NOLOCK) 
			INNER JOIN LLFBPS.dbo.ProductCategoryJUNC WITH (NOLOCK) 
				ON LLFBPS.dbo.ProductCategory.ProductCategoryID = LLFBPS.dbo.ProductCategoryJUNC.ProductCategoryID  
		WHERE LLFBPS.dbo.ProductCategory.Active = 1 
			AND LLFBPS.dbo.ProductCategoryJUNC.ParentCategoryID = 0  
			AND LLFBPS.dbo.ProductCategory.BrandID = @BrandID 
			AND LEFT(LLFBPS.dbo.ProductCategory.Name,9)  = 'AromaGlow') ,0) AS ParentCategoryID
	, NULL AS Roles
	, 70 AS Sequence
	, '~/Consumer/AromaGlowHowtoUse.aspx' AS url 
	, '' AS thumbnail 

/* Append static Where to buy items to submenu */
/***********************************************/
UNION 

/* Append static How to items to submenu */
/***********************************************/

SELECT 
	-4 as ProductCategoryID, 
	'Where to Buy' AS [Name], 
	'Where to Buy' AS Description
	, ISNULL((SELECT LLFBPS.dbo.ProductCategory.ProductCategoryID 
		FROM LLFBPS.dbo.ProductCategory WITH (NOLOCK) 
			INNER JOIN LLFBPS.dbo.ProductCategoryJUNC WITH (NOLOCK) 
				ON LLFBPS.dbo.ProductCategory.ProductCategoryID = LLFBPS.dbo.ProductCategoryJUNC.ProductCategoryID  
		WHERE LLFBPS.dbo.ProductCategory.Active = 1 
			AND LLFBPS.dbo.ProductCategoryJUNC.ParentCategoryID = 0  
			AND LLFBPS.dbo.ProductCategory.BrandID = @BrandID 
			AND LEFT(LLFBPS.dbo.ProductCategory.Name,4)  = 'Tiki') ,0) AS ParentCategoryID
	, NULL AS Roles
	, 80 AS Sequence
	, '~/Consumer/TikiWheretoBuy.aspx' AS url 
	, '' AS thumbnail 

UNION 

SELECT	-5 AS ProductCategoryID
	, 'Where to Buy' AS [Name]
	, 'Where to Buy' AS Description
	,  ISNULL((SELECT LLFBPS.dbo.ProductCategory.ProductCategoryID 
		FROM LLFBPS.dbo.ProductCategory WITH (NOLOCK) 
			INNER JOIN LLFBPS.dbo.ProductCategoryJUNC WITH (NOLOCK) 
				ON LLFBPS.dbo.ProductCategory.ProductCategoryID = LLFBPS.dbo.ProductCategoryJUNC.ProductCategoryID  
		WHERE LLFBPS.dbo.ProductCategory.Active = 1 
			AND LLFBPS.dbo.ProductCategoryJUNC.ParentCategoryID = 0  
			AND LLFBPS.dbo.ProductCategory.BrandID = @BrandID 
			AND LEFT(LLFBPS.dbo.ProductCategory.Name,9)  = 'Lamplight') ,0) AS ParentCategoryID
	, NULL AS Roles
	, 90 AS Sequence
	, '~/Consumer/LamplightWheretoBuy.aspx' AS url 
	, '' AS thumbnail 

UNION 

SELECT 
	-6 as ProductCategoryID
	, 'Where to Buy' AS [Name]
	, 'Where to Buy' AS Description
	, ISNULL((SELECT LLFBPS.dbo.ProductCategory.ProductCategoryID 
		FROM LLFBPS.dbo.ProductCategory WITH (NOLOCK) 
		INNER JOIN LLFBPS.dbo.ProductCategoryJUNC WITH (NOLOCK) 
			ON LLFBPS.dbo.ProductCategory.ProductCategoryID = LLFBPS.dbo.ProductCategoryJUNC.ProductCategoryID  
		WHERE LLFBPS.dbo.ProductCategory.Active = 1 
			AND LLFBPS.dbo.ProductCategoryJUNC.ParentCategoryID = 0  
			AND LLFBPS.dbo.ProductCategory.BrandID = @BrandID 
			AND Left(LLFBPS.dbo.ProductCategory.Name,9)  = 'AromaGlow') ,0) AS ParentCategoryID
	, NULL AS Roles
	, 100 AS Sequence
	, '~/Consumer/AromaGlowWheretoBuy.aspx' AS url 
	,  '' AS thumbnail 

/***********************************************/

/* Ecommerce Products */
UNION 
-- Product Temp Table goes here...
SELECT	ProductCategoryID
	, [Name]
	, [Description]
	, ParentCategoryID
	, Roles
	, [Sequence]
	, URL
	, thumbnail 
FROM @Products
ORDER BY sequence, name, productcategoryid 
END

GO



/*************************************************************************************************************************
	Add a Manager Schema to LLFWebsite
 *************************************************************************************************************************/

/****** Object:  Schema [Manager]    Script Date: 03/04/2011 14:39:39 ******/
IF  EXISTS (SELECT * FROM sys.schemas WHERE name = N'Manager')
DROP SCHEMA [Manager]
GO

USE [LLFWebSite]
GO

/****** Object:  Schema [Manager]    Script Date: 03/04/2011 14:39:39 ******/
IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = 'Manager')
BEGIN    
	-- The schema must be run in its own batch!    
	EXEC( 'CREATE SCHEMA [Manager] AUTHORIZATION [dbo]' );
END
GO

/*************************************************************************************************************************
	Updated GetAllCategorizationSequence to include Sequence information
 *************************************************************************************************************************/

USE [LLFWebSite]
GO

/****** Object:  StoredProcedure [Manager].[GetAllCategorizationSequence]    Script Date: 03/04/2011 14:40:44 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[Manager].[GetAllCategorizationSequence]') AND type in (N'P', N'PC'))
DROP PROCEDURE [Manager].[GetAllCategorizationSequence]
GO

USE [LLFWebSite]
GO

/****** Object:  StoredProcedure [Manager].[GetAllCategorizationSequence]    Script Date: 03/04/2011 14:40:44 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		Corey Perrymond
-- Create date: 3/22/10
-- Description:	Get All Marketing Series
-- =============================================
CREATE PROCEDURE [Manager].[GetAllCategorizationSequence] --1
	-- Add the parameters for the stored procedure here
	@BrandID INT = 2

AS
BEGIN
	SET NOCOUNT ON;

	/* Get top categories */
	SELECT DISTINCT TOP (100) PERCENT 0 AS ProductSeriesID
		, 0 AS ProductCategoryID
		, ProductCategoryID AS ParentCategoryID
		, Name AS Categorization
		, (convert(varchar,ISNULL(Sequence,0)) + '.' + 
			convert(varchar,ProductCategoryID) + '.' +
			convert(varchar,0) + '.' + -- set default 
			convert(varchar,0) + '.' + -- set default 
			convert(varchar,0) + '.' + -- set default 
			convert(varchar,0)
		 ) AS SortOrder
	                         
	FROM	LLFBPS.dbo.vProductCategories AS P
	WHERE   Active = 1 
			AND ParentCategoryID = 0
			AND productCategoryID > 0 
			AND BrandID = @BrandID

	UNION 

	/* Get Sub Categories */
	SELECT DISTINCT TOP (100) PERCENT 
		0 as ProductSeriesID
		, ProductCategoryID
		, ParentCategoryID
		, ParentCategoryName + ' >  ' + Name AS Categorization
		, (convert(varchar,ISNULL(ParentSequence,0)) + '.' + 
			convert(varchar,ISNULL(ParentCategoryID,0)) + '.' +
			convert(varchar,ISNULL(Sequence,0)) + '.' +
			convert(varchar,ProductCategoryID) + '.' +
			convert(varchar,0) + '.' + -- set default 
			convert(varchar,0)
		) as SortOrder
	                         
	FROM	LLFBPS.dbo.vProductCategories AS P
	WHERE   Active = 1 
			AND productCategoryID > 0 
			AND ParentCategoryID > 0 
			AND BrandID = @BrandID
	
	ORDER BY SortOrder

END

GO

/*************************************************************************************************************************
	Add GetProductListingsBySequence Procedure
 *************************************************************************************************************************/
USE [LLFWebSite]
GO

/****** Object:  StoredProcedure [Manager].[GetProductListingsBySequence]    Script Date: 03/04/2011 14:41:51 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[Manager].[GetProductListingsBySequence]') AND type in (N'P', N'PC'))
DROP PROCEDURE [Manager].[GetProductListingsBySequence]
GO

USE [LLFWebSite]
GO

/****** Object:  StoredProcedure [Manager].[GetProductListingsBySequence]    Script Date: 03/04/2011 14:41:51 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [Manager].[GetProductListingsBySequence]-- 96,4
	-- Add the parameters for the stored procedure here
	--@ProductSeriesID INT, 
	@ProductCategoryID INT,
	@BrandID INT = 2

AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure her

/* Begin Normalize Sequence before returning results if needed */
IF EXISTS(SELECT COUNT(*) FROM LLFBPS.dbo.vproducts WITH (NOLOCK)
WHERE EcommerceCategoryID = @ProductCategoryID --AND productseriesid = @ProductSeriesID
		AND active = 1
		AND BrandID = @BrandID 
GROUP BY sequence HAVING COUNT(*) > 1) 
BEGIN 
	-- initialized sequence 
	DECLARE @Sequence INT 
	DECLARE @ProductID INT 
	DECLARE @Products TABLE
	(
	ProductID INT,
	Sequence INT 
	)

	-- set first sequence value 
	SET @Sequence = 1 

	-- Load current order 
	INSERT INTO @Products
	SELECT ProductID, ISNULL(Sequence,99) as Sequence
	FROM	LLFBPS.dbo.vProducts WITH (NOLOCK) 
	WHERE   Active = 1 /*AND ProductSeriesID = @ProductSeriesID */ 
			AND EcommerceCategoryID = @ProductCategoryID
			AND BrandID = @BrandID
	ORDER BY Sequence

	-- Now normalize AND sequence all items before return results. 
	WHILE EXISTS (SELECT ProductID FROM @Products ) 
	BEGIN
		-- pulled the first item 
		SET @ProductID = (SELECT TOP 1 ProductID FROM @Products ORDER BY Sequence) 

		-- Update Product's sequence to unique value 
		UPDATE LLFBPS.dbo.Product SET Sequence = @Sequence WHERE productID = @ProductID

		-- remove AND move to next item 
		DELETE @Products WHERE ProductID = @ProductID 

		SET @Sequence = @Sequence + 1 

	END -- end loop 
END  
/* End Normalize Sequence before returning results if needed */


	 /* Get Sub Categories */
	SELECT ProductID, [Name] + '(' + ProductModelNumber+')' AS [Name]
			, ImagePath,ISNULL(Sequence,0) AS Sequence
	FROM	LLFBPS.dbo.vProducts WITH (NOLOCK) 
	WHERE   Active = 1 
			AND EcommerceCategoryID = @ProductCategoryID
			AND BrandID = @BrandID
	ORDER BY Sequence
END

GO


/*************************************************************************************************************************
	Add [UpdateProductSequence] Procedure
 *************************************************************************************************************************/
USE [LLFWebSite]
GO

/****** Object:  StoredProcedure [Manager].[UpdateProductSequence]    Script Date: 03/04/2011 14:42:24 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[Manager].[UpdateProductSequence]') AND type in (N'P', N'PC'))
DROP PROCEDURE [Manager].[UpdateProductSequence]
GO

USE [LLFWebSite]
GO

/****** Object:  StoredProcedure [Manager].[UpdateProductSequence]    Script Date: 03/04/2011 14:42:24 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		Corey Perrymond
-- Create date: 3/22/10
-- Description:	UPdate Product series sequence
-- =============================================

CREATE PROCEDURE [Manager].[UpdateProductSequence]--1
	-- Add the parameters for the stored procedure here
	@ProductID INT,
	@Sequence INT,
	-- Dummy field required by reorder control 
	@Name varchar(50) = '',
	@ImagePath varchar(150)= '' 


AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure her
 /* Get Sub Categories */
UPDATE LLFbps.dbo.Product SET Sequence = @Sequence WHERE productID = @ProductID

END

GO

USE [LLFBPS]
GO

/****** Object:  StoredProcedure [dbo].[GetProductSequence]    Script Date: 03/17/2011 16:27:50 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[GetProductSequence]') AND type in (N'P', N'PC'))
DROP PROCEDURE [dbo].[GetProductSequence]
GO

USE [LLFBPS]
GO

/****** Object:  StoredProcedure [dbo].[GetProductSequence]    Script Date: 03/17/2011 16:27:50 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		Corey Perrymond
-- Create date: 
-- Description:	Returns a Sequence of Products
-- =============================================
CREATE PROCEDURE [dbo].[GetProductSequence] 
	-- Add the parameters for the stored procedure here
	@BrandID int = 0
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
	SELECT 
		ProductID AS ProductCategoryID
		, ProductModelNumber AS [Name] 
		, ISNULL([Name],ProductModelNumber) AS Description
		, EcommerceCategoryID AS ParentCategoryID
		, NULL AS Roles
		, 200 + Sequence AS [Sequence]
		, '~/product/'+
			CONVERT(varchar(20),ProductID)+'/'+
			[dbo].[ConvertStringToHTMLPageName]([Name])+
			'.aspx' AS url 
		, ImagePath_small AS thumbnail 
	FROM LLFBPS.dbo.vProducts 
	WHERE ISNULL(EcommerceCategoryID,0) > 0
		AND LLFBPS.dbo.vProducts.Active = 1 --IsSellable = 1
		AND EcommerceCategoryID IN (
			SELECT LLFBPS.dbo.ProductCategory.ProductCategoryID 
			FROM LLFBPS.dbo.ProductCategory INNER JOIN
				LLFBPS.dbo.ProductCategoryJUNC ON 
				LLFBPS.dbo.ProductCategory.ProductCategoryID = 
				LLFBPS.dbo.ProductCategoryJUNC.ProductCategoryID 

			WHERE LLFBPS.dbo.ProductCategoryJUNC.ParentCategoryID > 0 
				AND LLFBPS.dbo.ProductCategory.productcategoryid > 0 
				AND LLFBPS.dbo.ProductCategory.active = 1
				AND LLFBPS.dbo.ProductCategory.BrandID = @BrandID 
				AND LLFBPS.dbo.ProductCategory.IsEcommerceCategory = 1) 
	ORDER BY [Sequence]

END


GO

