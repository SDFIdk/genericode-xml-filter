<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet
    version="1.0"
    xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:gc="http://docs.oasis-open.org/codelist/ns/genericode/1.0/"
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
    exclude-result-prefixes="gc">
    
    <!-- XSLT 1.0 is used, as this transformation is written to be used in LibreOffice, which uses the libxslt library. -->

    <xsl:output
        method="xml"
        version="1.0"
        encoding="UTF-8"
        indent="yes" />
        
    <!-- This stylesheet will create sheets with the following names -->
    <xsl:param
        name="sheetNameIdentification"
        select="'Identification'" />
    <xsl:param
        name="sheetNameColumnSet"
        select="'Columns'" />
    <xsl:param
        name="sheetNameSimpleCodeList"
        select="'Values'" />

    <xsl:template match="/">
        <office:document-content
            xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
            office:version="1.3">

            <office:body>
                <office:spreadsheet>
                    <xsl:apply-templates select="gc:CodeList" />
                </office:spreadsheet>
            </office:body>
        </office:document-content>
    </xsl:template>

    <xsl:template match="gc:CodeList">
        <xsl:apply-templates select="Identification" />

        <xsl:apply-templates select="ColumnSet" />

        <table:table>
            <xsl:attribute name="table:name">
                <xsl:value-of select="$sheetNameSimpleCodeList" />
            </xsl:attribute>
            <!-- Write header row with id-attributes of the columns as values. -->
            <table:table-row>
                <xsl:for-each select="ColumnSet/Column">
                    <table:table-cell office:value-type="string">
                        <text:p>
                            <xsl:value-of select="@Id" />
                        </text:p>
                    </table:table-cell>
                </xsl:for-each>
            </table:table-row>
        
            <!-- Write the data in the following rows. -->
            <xsl:for-each select="SimpleCodeList/Row">
                <xsl:variable
                    name="rowPosition"
                    select="position()" />
                <table:table-row>
                    <xsl:for-each select="Value">
                        <xsl:variable
                            name="valuePosition"
                            select="position()" />
                        <!-- Find the id attribute of the column that is in the same position as the current value -->
                        <xsl:variable
                            name="correspondingColumnId"
                            select="../../../ColumnSet/Column[position() = $valuePosition]/@Id" />
                        <xsl:variable
                            name="valueColumnRef"
                            select="@ColumnRef" />
                        <!-- Only work with genericode files in which the order of the values is the same as the order of the declared columns.
                        If this test would not be present and the order was different, the content of the cell would be placed in the wrong column.
                        The transformation is terminated, no attempt is done to place the value in the correct column. -->
                        <xsl:choose>
                            <xsl:when test="$valueColumnRef = $correspondingColumnId">
                                <table:table-cell
                                    office:value-type="string"
                                    calcext:value-type="string">
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

    <xsl:template match="Identification">
        <table:table>
            <xsl:attribute name="table:name">
                <xsl:value-of select="$sheetNameIdentification" />
            </xsl:attribute>
            <xsl:for-each select="*">
                <xsl:choose>
                    <xsl:when test="count(*) = 0">
                        <xsl:call-template name="outputIdentificationRow">
                            <xsl:with-param
                                name="nameInTable"
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
                            so no need to go deeper. -->
                            <xsl:call-template name="outputIdentificationRow">
                                <xsl:with-param
                                    name="nameInTable"
                                    select="concat($elementName, '/', name())" />
                            </xsl:call-template>
                        </xsl:for-each>
                    </xsl:when>
                </xsl:choose>
            </xsl:for-each>
        </table:table>
    </xsl:template>

    <xsl:template name="outputIdentificationRow">
        <xsl:param name="nameInTable" />
        <table:table-row>
            <table:table-cell office:value-type="string">
                <text:p>
                    <xsl:value-of select="$nameInTable" />
                </text:p>
            </table:table-cell>
            <table:table-cell office:value-type="string">
                <text:p>
                    <xsl:value-of select="text()" />
                </text:p>
            </table:table-cell>
            <!-- Attributes, if present, are written in the third column -->
            <table:table-cell office:value-type="string">
                <text:p>
                    <xsl:call-template name="string-join-attributes" />
                </text:p>
            </table:table-cell>
        </table:table-row>
    </xsl:template>

    <xsl:template name="string-join-attributes">
        <xsl:if test="count(@*) > 0">
            <!-- Attributes are written in the style att1=value1, att2=value2, ... -->
            <!-- With XSLT 2 or later, this could be written more simple as
            string-join(concat(name(), '=', .), ', ') -->
            <xsl:for-each select="@*">
                <xsl:if test="position() > 1">
                    <xsl:value-of select="', '" />
                </xsl:if>
                <xsl:value-of select="concat(name(), '=', .)" />
            </xsl:for-each>
        </xsl:if>
    </xsl:template>

    <xsl:template match="ColumnSet">
        <table:table>
            <xsl:attribute name="table:name">
                <xsl:value-of select="$sheetNameColumnSet" />
            </xsl:attribute>

            <!-- Header row -->
            <table:table-row>
                <table:table-cell>
                    <text:p>
                        <xsl:value-of select="'Id'" />
                    </text:p>
                </table:table-cell>
                <table:table-cell>
                    <text:p>
                        <xsl:value-of select="'Use'" />
                    </text:p>
                </table:table-cell>
                <table:table-cell>
                    <text:p>
                        <xsl:value-of select="'LongName (da)'" />
                    </text:p>
                </table:table-cell>
                <table:table-cell>
                    <text:p>
                        <xsl:value-of select="'Description (da)'" />
                    </text:p>
                </table:table-cell>
                <table:table-cell>
                    <text:p>
                        <xsl:value-of select="'LongName (en)'" />
                    </text:p>
                </table:table-cell>
                <table:table-cell>
                    <text:p>
                        <xsl:value-of select="'Description (en)'" />
                    </text:p>
                </table:table-cell>
                <table:table-cell>
                    <text:p>
                        <xsl:value-of select="'Data type'" />
                    </text:p>
                </table:table-cell>
                <table:table-cell>
                    <text:p>
                        <xsl:value-of select="'Language of data'" />
                    </text:p>
                </table:table-cell>
                <table:table-cell>
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
                            <xsl:value-of select="LongName[@xml:lang = 'da']" />
                        </text:p>
                    </table:table-cell>
                    <table:table-cell>
                        <text:p>
                            <xsl:value-of select="Annotation/Description/dc:description[@xml:lang = 'da']" />
                        </text:p>
                    </table:table-cell>
                    <table:table-cell>
                        <text:p>
                            <xsl:value-of select="LongName[@xml:lang = 'en']" />
                        </text:p>
                    </table:table-cell>
                    <table:table-cell>
                        <text:p>
                            <xsl:value-of select="Annotation/Description/dc:description[@xml:lang = 'en']" />
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

</xsl:stylesheet>