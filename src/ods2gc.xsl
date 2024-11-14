<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet
    version="1.0"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:gc="http://docs.oasis-open.org/codelist/ns/genericode/1.0/"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:str="http://exslt.org/strings"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    exclude-result-prefixes="office table text"
    extension-element-prefixes="str">

    <xsl:output
        method="xml"
        version="1.0"
        encoding="UTF-8"
        indent="yes" />
        
    <!-- This stylesheet assumes sheets with the following names -->
    <xsl:variable
        name="sheetNameIdentification"
        select="'Identification'" />
    <xsl:variable
        name="sheetNameMetadata"
        select="'Additional code list metadata'" />
    <xsl:variable
        name="sheetNameColumnSet"
        select="'Columns'" />
    <xsl:variable
        name="sheetNameSimpleCodeList"
        select="'Values'" />
        
    <!-- Create a lookup key of the first table row in the Values sheet with 
    - index: the position of the cell within the table row
    - value: the table cell
    The position() function cannot be used in this case, as it will always return 1.
    In XSLT 1.0 stylesheets, it is an error for the value of either the use attribute or the match attribute to contain a variable reference,
    see also https://www.w3.org/TR/xslt-10/#key,
    therefore, the name of the sheet is hardcoded -->
    <xsl:key
        name="valuesTableColumnPositionKey"
        match="table:table[@table:name = 'Values']/table:table-row[position() = 1]/table:table-cell"
        use="count(preceding-sibling::table:table-cell) + 1" />

    <xsl:template match="/">
        <gc:CodeList xmlns:gc="http://docs.oasis-open.org/codelist/ns/genericode/1.0/">
            <xsl:apply-templates
                select="office:document/office:body/office:spreadsheet/table:table[@table:name = $sheetNameMetadata]"
                mode="codeListMetadata" />
            <xsl:apply-templates
                select="office:document/office:body/office:spreadsheet/table:table[@table:name = $sheetNameIdentification]"
                mode="identification" />
            <xsl:apply-templates
                select="office:document/office:body/office:spreadsheet/table:table[@table:name = $sheetNameColumnSet]"
                mode="columnset" />
            <xsl:apply-templates
                select="office:document/office:body/office:spreadsheet/table:table[@table:name = $sheetNameSimpleCodeList]"
                mode="values" />
        </gc:CodeList>
    </xsl:template>
    
    <!-- Use the mode attribute and not a reference to a $sheetNameXXX variable in the match attribute,
    as in XSLT 1.0 stylesheets, it is an error for the value of the match attribute to contain a variable reference,
    see also https://www.w3.org/TR/xslt-10/#section-Defining-Template-Rules -->
    <xsl:template
        match="table:table"
        mode="codeListMetadata">
        <!-- Do not convert if sheet does not contain text -->
        <xsl:if test="count(table:table-row/table:table-cell/text:p) > 0">
            <Annotation>
                <Description>
                    <xsl:for-each select="table:table-row">
                        <xsl:call-template name="convertTableRowToXmlElement" />
                    </xsl:for-each>
                </Description>
            </Annotation>
        </xsl:if>
    </xsl:template>

    <!-- Context table:table-row -->
    <xsl:template name="convertTableRowToXmlElement">
        <xsl:if test="count(table:table-cell/@table:number-columns-repeated) > 0">
            <xsl:message terminate="yes">
                <xsl:text>This stylesheet does not take into account adjacent cells with the same content in sheets </xsl:text>
                <xsl:value-of select="$sheetNameIdentification" />
                <xsl:text> and </xsl:text>
                <xsl:value-of select="$sheetNameMetadata" />
                <xsl:text>.</xsl:text>
            </xsl:message>
        </xsl:if>
        <!-- Assume that the first cell contains only one paragraph -->
        <xsl:variable
            name="textFirstCell"
            select="table:table-cell[1]/text:p[1]" />
        <xsl:choose>
            <!-- no attributes; no child element, e.g. Version -->
            <xsl:when test="not(contains($textFirstCell, ' ')) and not(contains($textFirstCell, '/'))">
                <!-- Curly brackets are needed in the specification of the name of the element!
                They inform the XSLT processor that the contents need to be treated as XPath. -->
                <xsl:element name="{$textFirstCell}">
                    <xsl:apply-templates
                        select="table:table-cell[2]"
                        mode="joinTableCellParagraphs" />
                </xsl:element>
            </xsl:when>
            <!-- attributes; no child element, e.g. AlternateFormatLocationUri MimeType=text/csv -->
            <xsl:when test="contains($textFirstCell, ' ') and not(contains(substring-before($textFirstCell, ' '), '/'))">
                <!-- attributes -->
                <xsl:element name="{substring-before($textFirstCell, ' ')}">
                    <xsl:call-template name="tokenizeAttributestringAndCreateAttributes">
                        <xsl:with-param
                            name="attributestring"
                            select="substring-after($textFirstCell, ' ')" />
                    </xsl:call-template>
                    <xsl:apply-templates
                        select="table:table-cell[2]"
                        mode="joinTableCellParagraphs" />
                </xsl:element>
            </xsl:when>
            <!-- no attributes; child element, e.g. Agency/LongName -->
            <xsl:when test="not(contains($textFirstCell, ' ')) and contains($textFirstCell, '/')">
                <xsl:element name="{substring-after($textFirstCell, '/')}">
                    <xsl:apply-templates
                        select="table:table-cell[2]"
                        mode="joinTableCellParagraphs" />
                </xsl:element>
            </xsl:when>
            <!-- attributes; child element, e.g. Agency/LongName xml:lang=da -->
            <xsl:when test="contains($textFirstCell, ' ') and contains(substring-before($textFirstCell, ' '), '/')">
                <!-- attributes -->
                <xsl:element name="{substring-after(substring-before($textFirstCell, ' '), '/')}">
                    <xsl:call-template name="tokenizeAttributestringAndCreateAttributes">
                        <xsl:with-param
                            name="attributestring"
                            select="substring-after($textFirstCell, ' ')" />
                    </xsl:call-template>
                    <xsl:apply-templates
                        select="table:table-cell[2]"
                        mode="joinTableCellParagraphs" />
                </xsl:element>
            </xsl:when>
        </xsl:choose>
    </xsl:template>

    <xsl:template
        match="table:table"
        mode="identification">
        <Identification>
            <xsl:for-each select="table:table-row[not(starts-with(table:table-cell[position() = 1]/text:p, 'Agency/'))]">
                <xsl:call-template name="convertTableRowToXmlElement" />
            </xsl:for-each>
            <Agency>
                <xsl:for-each select="table:table-row[starts-with(table:table-cell[position() = 1]/text:p, 'Agency/')]">
                    <xsl:call-template name="convertTableRowToXmlElement" />
                </xsl:for-each>
            </Agency>
        </Identification>
    </xsl:template>

    <xsl:template name="tokenizeAttributestringAndCreateAttributes">
        <xsl:param name="attributestring" />
        <xsl:if test="string-length($attributestring) > 0">
            <xsl:for-each select="str:tokenize($attributestring, ',')">
                <xsl:attribute name="{substring-before(., '=')}">
                    <xsl:value-of select="substring-after(., '=')" />
                </xsl:attribute>
            </xsl:for-each>
        </xsl:if>
    </xsl:template>

    <xsl:template
        match="table:table"
        mode="columnset">

        <xsl:variable
            name="positionId"
            select="count(table:table-row[position() = 1]/table:table-cell[text:p/text() = '@Id']/preceding-sibling::table:table-cell) + 1" />
        <xsl:variable
            name="positionUse"
            select="count(table:table-row[position() = 1]/table:table-cell[text:p/text() = '@Use']/preceding-sibling::table:table-cell) + 1" />
        <xsl:variable
            name="positionDataType"
            select="count(table:table-row[position() = 1]/table:table-cell[text:p/text() = 'Data/@Type']/preceding-sibling::table:table-cell) + 1" />
        <xsl:variable
            name="positionDataLang"
            select="count(table:table-row[position() = 1]/table:table-cell[text:p/text() = 'Data/@Lang']/preceding-sibling::table:table-cell) + 1" />
        <xsl:variable
            name="positionDctermsDescription"
            select="count(table:table-row[position() = 1]/table:table-cell[text:p/text() = 'Annotation/Description/dcterms:description']/preceding-sibling::table:table-cell) + 1" />
        <xsl:variable
            name="positionKey"
            select="count(table:table-row[position() = 1]/table:table-cell[text:p/text() = 'Key']/preceding-sibling::table:table-cell) + 1" />

        <ColumnSet>
            <xsl:for-each select="table:table-row[position() > 1]">
                <xsl:if test="count(table:table-cell/@table:number-columns-repeated) > 0">
                    <xsl:message terminate="yes">
                        <xsl:text>This stylesheet does not take into account adjacent cells with the same content in sheet </xsl:text>
                        <xsl:value-of select="$sheetNameColumnSet" />
                        <xsl:text>.</xsl:text>
                    </xsl:message>
                </xsl:if>

                <Column>
                    <xsl:attribute name="Id">
                        <xsl:value-of select="table:table-cell[position() = $positionId]/text:p" />
                    </xsl:attribute>
                    <xsl:attribute name="Use">
                        <xsl:value-of select="table:table-cell[position() = $positionUse]/text:p" />
                    </xsl:attribute>
                    <Annotation>
                        <Description>
                            <dcterms:description>
                                <xsl:value-of select="table:table-cell[position() = $positionDctermsDescription]/text:p" />
                            </dcterms:description>
                        </Description>
                    </Annotation>
                    <ShortName>
                        <!-- Use id attribute as short name -->
                        <xsl:value-of select="table:table-cell[position() = $positionId]/text:p" />
                    </ShortName>
                    <Data>
                        <xsl:attribute name="Type">
                            <xsl:value-of select="table:table-cell[position() = $positionDataType]/text:p" />
                        </xsl:attribute>
                        <xsl:if test="string-length(table:table-cell[position() = $positionDataLang]/text:p) > 0">
                            <xsl:attribute name="Lang">
                                <xsl:value-of select="table:table-cell[position() = $positionDataLang]/text:p" />
                            </xsl:attribute>
                        </xsl:if>
                    </Data>
                </Column>
            </xsl:for-each>

            <xsl:for-each select="table:table-row[position() > 1]">
                <xsl:if test="string-length(table:table-cell[position() = $positionKey]/text:p) > 0">
                    <Key>
                        <xsl:attribute name="Id">
                            <xsl:value-of select="table:table-cell[position() = $positionKey]/text:p" />
                        </xsl:attribute>
                        <ShortName>
                            <xsl:value-of select="table:table-cell[position() = $positionKey]/text:p" />
                        </ShortName>
                        <ColumnRef>
                            <xsl:attribute name="Ref">
                                <xsl:value-of select="table:table-cell[position() = $positionId]/text:p" />
                            </xsl:attribute>
                        </ColumnRef>
                    </Key>
                </xsl:if>
            </xsl:for-each>
        </ColumnSet>
    </xsl:template>

    <xsl:template
        match="table:table"
        mode="values">
        <!-- The header row is used to create a lookup key, see above;
        process the actual data in the rest of the rows -->
        <SimpleCodeList>
            <xsl:for-each select="table:table-row[position() > 1]">
                <!-- Do not take into account empty rows (rows with only whitespace will be converted though) -->
                <xsl:if test="count(table:table-cell/text:p) > 0">
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
                </xsl:if>
            </xsl:for-each>
        </SimpleCodeList>
    </xsl:template>

    <!-- Context table:table-cell -->
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
            	<xsl:value-of select="key('valuesTableColumnPositionKey', $columnPosition)/text:p" />
            </xsl:attribute>
            <!-- In this transformation, an undefined value (an empty string in the cell) (only applicable in optional columns)
            is always written as a Value element that does not contain a SimpleValue element. -->
            <xsl:if test="string-length(normalize-space(text:p)) > 0">
                <SimpleValue>
                    <xsl:apply-templates
                        select="."
                        mode="joinTableCellParagraphs" />
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

    <!-- A cell contains several paragraphs if the user has inserted a line break in it,
    see also https://help.libreoffice.org/latest/en-US/text/shared/guide/breaking_lines.html
    (tested with LibreOffice 24.8.0.3).
    This template is called for cells that potentially contain several paragraphs. -->
    <xsl:template
        match="table:table-cell"
        mode="joinTableCellParagraphs">
        <xsl:for-each select="text:p">
            <xsl:value-of select="text()" />
            <!-- Add a new line if more paragraphs follow -->
            <xsl:if test="not(position()=last())">
                <xsl:value-of select="'&#10;'" />
            </xsl:if>
        </xsl:for-each>
    </xsl:template>

</xsl:stylesheet>