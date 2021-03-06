DECLARE @Address      VARCHAR(100), 
        @City         VARCHAR(25), 
        @State        VARCHAR(25), 
        @PostalCode   VARCHAR(10), 
        @Country      VARCHAR(40), 
        @County       VARCHAR(40), 
        @GPSLatitude  NUMERIC(18, 9), 
        @GPSLongitude NUMERIC(18, 9), 
        @MapURL       VARCHAR(1024), 
        @AppID        VARCHAR(30), 
        @AppCode      VARCHAR(30) 

--Address can be any format upto and including city, state, postal, etc. example: @Address = '235 Aero Way NE Calgary, Alberta, Canada'
SET @Address = '235 Aero Way NE'
SET @City = 'Calgary'
SET @State = 'Alberta'
SET @AppID = ''
SET @AppCode = ''

DECLARE @URL varchar(MAX)
SET @URL = 'https://geocoder.api.here.com/6.2/geocode.xml?app_id=' + 
	COALESCE(@AppID, '') + '&app_code=' + 
	COALESCE(@AppCode, '') + '&searchtext=' +
	COALESCE(@Address, '') + ' ' +
	COALESCE(@City, '') + ' ' +
	COALESCE(@State, '')

SET @URL = REPLACE(@URL, ' ', '+')

DECLARE @Response varchar(8000)
DECLARE @XML xml
DECLARE @Obj int
DECLARE @Result int
DECLARE @HTTPStatus int
DECLARE @ErrorMsg varchar(MAX)

EXEC @Result = sp_OACreate 'MSXML2.ServerXMLHttp', @Obj OUT

BEGIN TRY
EXEC @Result = sp_OAMethod @Obj, 'open', NULL, 'GET', @URL, false
EXEC @Result = sp_OAMethod @Obj, 'setRequestHeader', NULL, 'Content-Type', 'application/x-www-form-urlencoded'
EXEC @Result = sp_OAMethod @Obj, send, NULL, ''
EXEC @Result = sp_OAGetProperty @Obj, 'status', @HTTPStatus OUT
EXEC @Result = sp_OAGetProperty @Obj, 'responseXML.xml', @Response OUT
END TRY
BEGIN CATCH	SET @ErrorMsg = ERROR_MESSAGE()END CATCH
EXEC @Result = sp_OADestroy @Obj

IF (@ErrorMsg IS NOT NULL) OR (@HTTPStatus <> 200)
BEGIN
	SET @ErrorMsg = 'Error Geocodeing: ' + ISNULL(@ErrorMsg, 'HTTP status code: ' + CAST(@HTTPStatus as varchar(10)))
	RAISERROR(@ErrorMsg, 16, 1, @HTTPStatus) RETURN END

SET @XML = CAST(@Response AS XML)
SET @GPSLatitude = @XML.value('declare namespace ns2="http://www.navteq.com/lbsp/Search-Search/4"; (/ns2:Search/Response/View/Result/Location/NavigationPosition/Latitude) [1]', 'numeric(18,9)')
SET @GPSLongitude =	@XML.value('declare namespace ns2="http://www.navteq.com/lbsp/Search-Search/4"; (/ns2:Search/Response/View/Result/Location/NavigationPosition/Longitude) [1]', 'numeric(18,9)')
SET @City =	@XML.value('declare namespace ns2="http://www.navteq.com/lbsp/Search-Search/4"; (/ns2:Search/Response/View/Result/Location/Address/City) [1]', 'varchar(40)')
SET @State = @XML.value('declare namespace ns2="http://www.navteq.com/lbsp/Search-Search/4"; (/ns2:Search/Response/View/Result/Location/Address/State) [1]', 'varchar(40)')
SET @PostalCode = @XML.value('declare namespace ns2="http://www.navteq.com/lbsp/Search-Search/4"; (/ns2:Search/Response/View/Result/Location/Address/PostalCode) [1]', 'varchar(20)')
SET @Country =	@XML.value('declare namespace ns2="http://www.navteq.com/lbsp/Search-Search/4"; (/ns2:Search/Response/View/Result/Location/Address/Country) [1]', 'varchar(40)')
SET @Address = ISNULL(@XML.value('declare namespace ns2="http://www.navteq.com/lbsp/Search-Search/4"; (/ns2:Search/Response/View/Result/Location/Address/HouseNumber) [1]', 'varchar(40)'), '???') + ' ' + 
			   ISNULL(@XML.value('declare namespace ns2="http://www.navteq.com/lbsp/Search-Search/4"; (/ns2:Search/Response/View/Result/Location/Address/Street) [1]', 'varchar(40)'), '???')

SELECT	@Address 'Address', @City 'City', @State 'State', @PostalCode 'Postal Code', @Country 'Country', @GPSLatitude 'Latitude', @GPSLongitude 'Longitude'
