<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet
    version="1.0"
    xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
    xmlns:gc="http://docs.oasis-open.org/codelist/ns/genericode/1.0/"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    exclude-result-prefixes="gc dcterms">

    <xsl:output
        method="xml"
        version="1.0"
        encoding="UTF-8"
        indent="yes" />
        
    <!-- This stylesheet will create sheets with the following names -->
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
        
    <!-- Create a lookup key with 
    - index: the position of the Column within the ColumnSet
    - value: the Column element itself -->
    <xsl:key
        name="columnPositionAndIdKey"
        match="ColumnSet/Column"
        use="count(preceding-sibling::Column) + 1" />

    <xsl:template match="/gc:CodeList">
        <office:document
            office:version="1.3"
            office:mimetype="application/vnd.oasis.opendocument.spreadsheet">

            <office:automatic-styles>
                <style:style
                    style:name="columnnormal"
                    style:family="table-column">
                    <style:table-column-properties style:column-width="2.5cm" />
                </style:style>
                <style:style
                    style:name="columnwide"
                    style:family="table-column">
                    <style:table-column-properties style:column-width="10cm" />
                </style:style>
                <style:style
                    style:name="columnverywide"
                    style:family="table-column">
                    <style:table-column-properties style:column-width="25cm" />
                </style:style>
                <style:style
                    style:name="cellheader"
                    style:family="table-cell">
                    <style:text-properties fo:font-weight="bold" />
                </style:style>
                <style:style
                    style:name="cellvalue"
                    style:family="table-cell">
                    <style:table-cell-properties fo:wrap-option="wrap" />
                    <style:text-properties fo:font-weight="normal" />
                </style:style>
            </office:automatic-styles>

            <office:body>
                <office:spreadsheet>
                    <xsl:apply-templates select="Identification" />
                    <xsl:apply-templates select="Annotation/Description" />
                    <xsl:apply-templates select="ColumnSet" />
                    <xsl:apply-templates select="SimpleCodeList" />
                </office:spreadsheet>

            </office:body>
        </office:document>
    </xsl:template>

    <xsl:template match="Identification">
        <table:table>
            <xsl:attribute name="table:name">
                <xsl:value-of select="$sheetNameIdentification" />
            </xsl:attribute>
            <table:table-column
                table:style-name="columnwide"
                table:default-cell-style-name="cellheader" />
            <table:table-column
                table:style-name="columnverywide"
                table:default-cell-style-name="cellvalue" />
            <xsl:for-each select="*">
                <xsl:choose>
                    <xsl:when test="count(*) = 0">
                        <xsl:call-template name="outputMetadataElementRow">
                            <xsl:with-param
                                name="nameInRow"
                                select="name()" />
                        </xsl:call-template>
                    </xsl:when>
                    <xsl:when test="count(*) > 0">
                        <!-- Store the name in a variable, as using
                        concat(../name(), '/', name())
                        for the parameter gives the following error with libxslt
                        XSLT-with-param: Failed to compile select expression 'concat(../name(), '/', name())' -->
                        <xsl:variable
                            name="elementName"
                            select="name()" />
                        <xsl:for-each select="*">
                            <!-- Only Identification/Agency has children, and those children are leaf elements,
                            so no need to go deeper. Otherwise, a better solution would be to create a 
                            recursive template. -->
                            <xsl:call-template name="outputMetadataElementRow">
                                <xsl:with-param
                                    name="nameInRow"
                                    select="concat($elementName, '/', name())" />
                            </xsl:call-template>
                        </xsl:for-each>
                    </xsl:when>
                </xsl:choose>
            </xsl:for-each>
        </table:table>
    </xsl:template>

    <xsl:template name="outputMetadataElementRow">
        <xsl:param name="nameInRow" />
        <table:table-row>
            <table:table-cell office:value-type="string">
                <text:p>
                    <xsl:value-of select="$nameInRow" />
                    <xsl:if test="count(@*) > 0">
                        <!-- Whitespace to separate element name from attributes -->
                        <xsl:value-of select="' '" />
                        <xsl:call-template name="stringJoinAttributes" />
                    </xsl:if>
                </text:p>
            </table:table-cell>
            <table:table-cell office:value-type="string">
                <text:p>
                    <xsl:value-of select="text()" />
                </text:p>
            </table:table-cell>
        </table:table-row>
    </xsl:template>

    <xsl:template name="stringJoinAttributes">
        <!-- Attributes are written in the style att1=value1,att2=value2, ... 
        Do not use double quotation marks att1="value1",att2="value2", ... 
        as LibreOffice may turn them into curly quotation marks! -->
        <xsl:for-each select="@*">
            <xsl:if test="position() > 1">
                <xsl:value-of select="','" />
            </xsl:if>
            <xsl:value-of select="concat(name(), '=', .)" />
        </xsl:for-each>
    </xsl:template>

    <xsl:template match="gc:CodeList/Annotation/Description">
        <table:table>
            <xsl:attribute name="table:name">
                <xsl:value-of select="$sheetNameMetadata" />
            </xsl:attribute>
            <table:table-column
                table:style-name="columnwide"
                table:default-cell-style-name="cellheader" />
            <table:table-column
                table:style-name="columnverywide"
                table:default-cell-style-name="cellvalue" />
            <xsl:for-each select="dcterms:*">
                <xsl:call-template name="outputMetadataElementRow">
                    <xsl:with-param
                        name="nameInRow"
                        select="name()" />
                </xsl:call-template>
            </xsl:for-each>
        </table:table>
    </xsl:template>

    <xsl:template match="ColumnSet">
        <table:table>
            <xsl:attribute name="table:name">
                <xsl:value-of select="$sheetNameColumnSet" />
            </xsl:attribute>
            <table:table-column
                table:number-columns-repeated="4"
                table:style-name="columnnormal"
                table:default-cell-style-name="cellvalue" />
            <table:table-column
                table:style-name="columnverywide"
                table:default-cell-style-name="cellvalue" />
            <table:table-column
                table:style-name="columnnormal"
                table:default-cell-style-name="cellvalue" />
            <!-- Header row -->
            <table:table-row>
                <table:table-cell table:style-name="cellheader">
                    <text:p>
                        <xsl:value-of select="'@Id'" />
                    </text:p>
                </table:table-cell>
                <table:table-cell table:style-name="cellheader">
                    <text:p>
                        <xsl:value-of select="'@Use'" />
                    </text:p>
                </table:table-cell>
                <!-- ShortName not present although it is mandatory: is set to same as @Id in ods2gc.xsl -->
                <table:table-cell table:style-name="cellheader">
                    <text:p>
                        <xsl:value-of select="'Data/@Type'" />
                    </text:p>
                </table:table-cell>
                <table:table-cell table:style-name="cellheader">
                    <text:p>
                        <xsl:value-of select="'Data/@Lang'" />
                    </text:p>
                </table:table-cell>
                <table:table-cell table:style-name="cellheader">
                    <text:p>
                        <xsl:value-of select="'Annotation/Description/dcterms:description'" />
                    </text:p>
                </table:table-cell>
                <table:table-cell table:style-name="cellheader">
                    <text:p>
                        <xsl:value-of select="'Key'" />
                    </text:p>
                </table:table-cell>
            </table:table-row>

            <xsl:for-each select="Column">
                <table:table-row>
                    <table:table-cell>
                        <text:p>
                            <xsl:value-of select="@Id" />
                        </text:p>
                    </table:table-cell>
                    <table:table-cell>
                        <text:p>
                            <xsl:value-of select="@Use" />
                        </text:p>
                    </table:table-cell>
                    <table:table-cell>
                        <text:p>
                            <xsl:value-of select="Data/@Type" />
                        </text:p>
                    </table:table-cell>
                    <table:table-cell>
                        <text:p>
                            <xsl:value-of select="Data/@Lang" />
                        </text:p>
                    </table:table-cell>
                    <table:table-cell>
                        <text:p>
                            <xsl:value-of select="Annotation/Description/dcterms:description" />
                        </text:p>
                    </table:table-cell>
                    <xsl:variable
                        name="columnId"
                        select="@Id" />
                    <table:table-cell>
                        <text:p>
                            <xsl:value-of select="../Key[ColumnRef/@Ref = $columnId]/@Id" />
                        </text:p>
                    </table:table-cell>
                </table:table-row>
            </xsl:for-each>
        </table:table>
    </xsl:template>

    <xsl:template match="SimpleCodeList">
        <table:table>
            <xsl:attribute name="table:name">
                <xsl:value-of select="$sheetNameSimpleCodeList" />
            </xsl:attribute>

            <table:table-column
                table:number-columns-repeated="{count(../ColumnSet/Column)}"
                table:style-name="columnnormal"
                table:default-cell-style-name="cellvalue" />
                
            <!-- Write header row with id-attributes of the columns as values. -->
            <table:table-row>
                <xsl:for-each select="../ColumnSet/Column">
                    <table:table-cell
                        office:value-type="string"
                        table:style-name="cellheader">
                        <text:p>
                            <xsl:value-of select="@Id" />
                        </text:p>
                    </table:table-cell>
                </xsl:for-each>
            </table:table-row>
        
            <!-- Write the data in the following rows. -->
            <xsl:for-each select="Row">
                <xsl:variable
                    name="rowPosition"
                    select="position()" />
                <table:table-row>
                    <xsl:for-each select="Value">
                        <xsl:variable
                            name="valuePosition"
                            select="position()" />
                        <xsl:variable
                            name="valueColumnRef"
                            select="@ColumnRef" />
                        <!-- Only work with genericode files in which the order of the values is the same as the order of the declared columns.
                        If this test would not be present and the order was different, the contents of the cell would be placed in the wrong column.
                        The transformation is terminated, no attempt is done to place the value in the correct column. -->
                        <xsl:choose>
                            <xsl:when test="$valueColumnRef = key('columnPositionAndIdKey', $valuePosition)/@Id">
                                <table:table-cell office:value-type="string">
                                    <text:p>
                                        <xsl:value-of select="SimpleValue" />
                                    </text:p>
                                </table:table-cell>
                            </xsl:when>
                            <xsl:otherwise>
                                <xsl:message terminate="yes">
                                    <xsl:value-of select="'The values in row '" />
                                    <xsl:value-of select="$rowPosition" />
                                    <xsl:value-of select="' are not in the same order as the columns in the column set, terminating here.'" />
                                </xsl:message>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:for-each>
                </table:table-row>
            </xsl:for-each>
        </table:table>
    </xsl:template>

</xsl:stylesheet>