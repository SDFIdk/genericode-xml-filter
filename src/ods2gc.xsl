<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet
    version="1.0"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:gc="http://docs.oasis-open.org/codelist/ns/genericode/1.0/"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:str="http://exslt.org/strings"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    exclude-result-prefixes="gc office table text"
    extension-element-prefixes="str">

    <xsl:output
        method="xml"
        version="1.0"
        encoding="UTF-8"
        indent="yes" />
        
    <!-- This stylesheet assumes sheets with the following names -->
    <xsl:param
        name="sheetNameIdentification"
        select="'Identification'" />
    <xsl:param
        name="sheetNameColumnSet"
        select="'Columns'" />
    <xsl:param
        name="sheetNameSimpleCodeList"
        select="'Values'" />
        
    <!-- Create a lookup key with 
    - index: the position of the cell within the table row
    - value: the table cell
    The position() function cannot be used in this case, as it will always return 1.
    In XSLT 1.0 stylesheets, it is an error for the value of either the use attribute or the match attribute to contain a variable reference,
    see also https://www.w3.org/TR/xslt-10/#key,
    therefore, the name of the sheet is hardcoded -->
    <xsl:key
        name="valuesColumnPositionAndNameKey"
        match="table:table[@table:name = 'Values']/table:table-row[position() = 1]/table:table-cell"
        use="count(preceding-sibling::table:table-cell) + 1" />

    <xsl:template match="/">
        <gc:CodeList xmlns:gc="http://docs.oasis-open.org/codelist/ns/genericode/1.0/">
            <!-- This stylesheet assumes that the ODS file has sheets with the following names.
            The values are hardcoded, as with XSLT 1.0, variables are not allowed in match,
            see also https://www.w3.org/TR/xslt-10/#section-Defining-Template-Rules -->
            <xsl:apply-templates select="//office:spreadsheet/table:table[@table:name = $sheetNameIdentification]" mode="identification" />
            <xsl:apply-templates select="//office:spreadsheet/table:table[@table:name = $sheetNameColumnSet]" mode="columnset" />
            <xsl:apply-templates select="//office:spreadsheet/table:table[@table:name = $sheetNameSimpleCodeList]" mode="values" />
        </gc:CodeList>
    </xsl:template>

    <!-- Use the mode attribute and not a reference to $sheetNameIdentification in the match attribute,
    as in XSLT 1.0 stylesheetss, it is an error for the value of the match attribute to contain a variable reference,
    see also https://www.w3.org/TR/xslt-10/#section-Defining-Template-Rules -->
    <xsl:template match="table:table" mode="identification">
        <Identification>
            <xsl:for-each select="table:table-row[not(starts-with(table:table-cell[position() = 1]/text:p, 'Agency/'))]">
                <!-- Curly brackets are needed in the specification of the name of the element!
                They inform the XSLT processor that the contents need to be treated as XPath. -->
                <xsl:element name="{table:table-cell[position() = 1]/text:p}">
                    <xsl:call-template name="tokenize-attributestring-and-create-attributes" />
                    <xsl:value-of select="table:table-cell[position() = 2]/text:p" />
                </xsl:element>
            </xsl:for-each>
            <Agency>
                <xsl:for-each select="table:table-row[starts-with(table:table-cell[position() = 1]/text:p, 'Agency/')]">
                    <xsl:element name="{substring-after(table:table-cell[position() = 1]/text:p, 'Agency/')}">
                        <xsl:call-template name="tokenize-attributestring-and-create-attributes" />
                        <xsl:value-of select="table:table-cell[position() = 2]/text:p" />
                    </xsl:element>
                </xsl:for-each>
            </Agency>
        </Identification>
    </xsl:template>

    <xsl:template name="tokenize-attributestring-and-create-attributes">
        <xsl:variable
            name="attributestring"
            select="table:table-cell[position() = 3]/text:p" />
        <xsl:if test="string-length($attributestring) > 0">
            <!-- Attributes, if present, are written in the third column, in the style att1=value1, att2=value2, ... -->
            <xsl:for-each select="str:tokenize($attributestring, ', ')">
                <xsl:attribute name="{substring-before(., '=')}">
                    <xsl:value-of select="substring-after(., '=')" />
                </xsl:attribute>
            </xsl:for-each>
        </xsl:if>
    </xsl:template>

    <xsl:template match="table:table" mode="columnset">
        <ColumnSet>
            <xsl:for-each select="table:table-row[position() > 1]">
                <!-- 
                 1: Id
                 2: Use
                 3: LongName (da)
                 4: Description (da)
                 5: LongName (en)
                 6: Description (en)
                 7: Data type
                 8: Language of data
                 9: Key
                 -->
                <Column>
                    <xsl:attribute name="Id">
                        <xsl:value-of select="table:table-cell[position() = 1]/text:p" />
                    </xsl:attribute>
                    <xsl:attribute name="Use">
                        <xsl:value-of select="table:table-cell[position() = 2]/text:p" />
                    </xsl:attribute>
                    <xsl:if test="string-length(table:table-cell[position() = 4]/text:p) > 0 or string-length(table:table-cell[position() = 6]/text:p) > 0">
                        <Annotation>
                            <Description>
                                <xsl:if test="string-length(table:table-cell[position() = 4]/text:p) > 0">
                                    <dc:description xml:lang="da">
                                        <xsl:value-of select="table:table-cell[position() = 4]/text:p" />
                                    </dc:description>
                                </xsl:if>
                                <xsl:if test="string-length(table:table-cell[position() = 6]/text:p) > 0">
                                    <dc:description xml:lang="en">
                                        <xsl:value-of select="table:table-cell[position() = 6]/text:p" />
                                    </dc:description>
                                </xsl:if>
                            </Description>
                        </Annotation>
                    </xsl:if>
                    <ShortName>
                        <!-- Use id attribute as shortname as well -->
                        <xsl:value-of select="table:table-cell[position() = 1]/text:p" />
                    </ShortName>
                    <LongName xml:lang="da">
                        <xsl:value-of select="table:table-cell[position() = 3]/text:p" />
                    </LongName>
                    <LongName xml:lang="en">
                        <xsl:value-of select="table:table-cell[position() = 5]/text:p" />
                    </LongName>
                    <Data>
                        <xsl:attribute name="Type">
                            <xsl:value-of select="table:table-cell[position() = 7]/text:p" />
                        </xsl:attribute>
                        <xsl:if test="string-length(table:table-cell[position() = 8]/text:p) > 0">
                            <xsl:attribute name="Lang">
                                <xsl:value-of select="table:table-cell[position() = 8]/text:p" />
                            </xsl:attribute>
                        </xsl:if>
                    </Data>
                </Column>
            </xsl:for-each>
            <xsl:for-each select="table:table-row[position() > 1]">
                <xsl:if test="string-length(table:table-cell[position() = 9]/text:p) > 0">
                    <Key>
                        <xsl:attribute name="Id">
                            <xsl:value-of select="table:table-cell[position() = 9]/text:p" />
                        </xsl:attribute>
                        <ShortName>
                            <xsl:value-of select="table:table-cell[position() = 9]/text:p" />
                        </ShortName>
                        <ColumnRef>
                            <xsl:attribute name="Ref">
                                <xsl:value-of select="table:table-cell[position() = 1]/text:p" />
                            </xsl:attribute>
                        </ColumnRef>
                    </Key>
                </xsl:if>
            </xsl:for-each>
        </ColumnSet>
    </xsl:template>
    
    <xsl:template match="table:table" mode="values">
        <!-- The header row is used to create a lookup key, see above;
        process the actual data in the rest of the rows -->
        <SimpleCodeList>
            <xsl:for-each select="table:table-row[position() > 1]">
                <Row>
                    <xsl:for-each select="table:table-cell">
                        <xsl:call-template name="writeValue">
                            <!-- Default value of table:number-columns-repeated is 1 according to the
                            OpenDocument Format specification, see also
                            https://docs.oasis-open.org/office/OpenDocument/v1.3/os/part3-schema/OpenDocument-v1.3-os-part3-schema.html#__RefHeading__1418526_253892949 -->
                            <xsl:with-param
                                name="columnPosition"
                                select="count(preceding-sibling::table:table-cell[not(@table:number-columns-repeated)]) + sum(preceding-sibling::table:table-cell/@table:number-columns-repeated) + 1" />
                            <xsl:with-param name="noOfRepetitions">
                                <xsl:choose>
                                    <xsl:when test="count(@table:number-columns-repeated) = 1">
                                        <xsl:value-of select="@table:number-columns-repeated" />
                                    </xsl:when>
                                    <xsl:otherwise>
                                        <xsl:value-of select="1" />
                                    </xsl:otherwise>
                                </xsl:choose>
                            </xsl:with-param>
                        </xsl:call-template>
                    </xsl:for-each>
                </Row>
            </xsl:for-each>
        </SimpleCodeList>
    </xsl:template>

    <xsl:template name="writeValue">
        <!-- number of the column the value is located in the spreadsheet visible in the GUI (column A is 1, column B is 2, etc.)  -->
        <xsl:param name="columnPosition" />
        <!-- number of times that the value is repeated in the successive columns  -->
        <xsl:param name="noOfRepetitions" />
        <!-- A Value element is always written, thus also if the cell contains an empty string.
        Following this convention makes it easier to transform to other formats.
        This is more strict than the genericode specification, which also allows rows that do not have a
        Value element corresponding to every column. -->
        <Value>
            <xsl:attribute name="ColumnRef">
            	<xsl:value-of select="key('valuesColumnPositionAndNameKey', $columnPosition)/text:p" />
            </xsl:attribute>
            <!-- In this transformation, an undefined value (an empty string in the cell) (only applicable in optional columns)
            is always written as a Value element that does not contain a SimpleValue element. -->
            <xsl:if test="string-length(normalize-space(text:p)) > 0">
                <SimpleValue>
                    <xsl:value-of select="text:p" />
                </SimpleValue>
            </xsl:if>
        </Value>
        <xsl:if test="$noOfRepetitions > 1">
            <!-- This is a recursive template -->
            <xsl:call-template name="writeValue">
                <xsl:with-param
                    name="columnPosition"
                    select="$columnPosition + 1" />
                <xsl:with-param
                    name="noOfRepetitions"
                    select="$noOfRepetitions - 1" />
            </xsl:call-template>
        </xsl:if>
    </xsl:template>

</xsl:stylesheet>