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
    
    <!-- XSLT 1.0 is used, as this transformation is written to be used in LibreOffice, which uses the libxslt library.
    libxslt supports most of the EXSLT extensions. For more informatation about EXSLT, see https://exslt.github.io/ -->

    <xsl:output
        method="xml"
        version="1.0"
        encoding="UTF-8"
        indent="yes" />

    <xsl:template match="/">
        <gc:CodeList xmlns:gc="http://docs.oasis-open.org/codelist/ns/genericode/1.0/">
            <!-- This stylesheet assumes that the ODS file has sheets with the following names.
            The values are hardcoded, as with XSLT 1.0, variables are not allowed in match,
            see also https://www.w3.org/TR/xslt-10/#section-Defining-Template-Rules -->
            <xsl:apply-templates select="//office:spreadsheet/table:table[@table:name = 'Identification']" />
            <xsl:apply-templates select="//office:spreadsheet/table:table[@table:name = 'Columns']" />
            <xsl:apply-templates select="//office:spreadsheet/table:table[@table:name = 'Values']" />
        </gc:CodeList>
    </xsl:template>

    <xsl:template match="table:table[@table:name = 'Identification']">
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

    <xsl:template match="table:table[@table:name = 'Columns']">
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

    <xsl:template match="table:table[@table:name = 'Values']">
        <!-- Ignore the header row and process the actual data in the rest of the rows -->
        <SimpleCodeList>
            <xsl:for-each select="table:table-row[position() > 1]">
                <Row>
                    <xsl:for-each select="table:table-cell">
                        <!-- Find the header name of the column this cell is located in by finding the
                        header cell in the same column as to the current cell being processed. -->
                        <xsl:variable
                            name="tableCellPosition"
                            select="position()" />
                        <!-- A Value element is always written, also if the cell contains an empty string.
                        Following this convention makes it easier to transform to other formats.
                        This is more strict than the genericode specification, which also allows rows that do not have a
                        Value element corresponding to every column.
                        In this transformation, an undefined value (an empty string in the cell) (only applicable in optional columns)
                        is always written as a Value element that does not contain a SimpleValue element. -->
                        <Value>
                            <xsl:attribute name="ColumnRef">
                                <!-- Having the column headers in a variable instead of checking the first row every time would be more efficient.
                                However, it is possible to do this in an elegant way with XSLT 1.0? -->
                                <xsl:value-of select="../../table:table-row[position() = 1]/table:table-cell[position() = $tableCellPosition]/text:p" />
                            </xsl:attribute>
                            <xsl:if test="string-length(normalize-space(text:p)) > 0">
                                <SimpleValue>
                                    <xsl:value-of select="text:p" />
                                </SimpleValue>
                            </xsl:if>
                        </Value>
                    </xsl:for-each>
                </Row>
            </xsl:for-each>
        </SimpleCodeList>
    </xsl:template>

</xsl:stylesheet>