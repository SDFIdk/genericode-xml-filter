# genericode-xml-filter

genericode-xml-filter is an [XML filter for LibreOffice](https://help.libreoffice.org/latest/en-US/text/shared/guide/xsltfilter.html) that converts genericode files to OpenDocument Spreadsheet documents. As the XML filter can be integrated into LibreOffice seamlessly, genericode files can be loaded and saved transparently.

## About the underlying standards and tools

### Genericode

[Genericode](https://docs.oasis-open.org/codelist/genericode/v1.0/genericode-v1.0.html), also known as Code List Representation, is “a single semantic model for code lists and accompanying XML serialization that is designed to IT-enable and standardize the publication of machine-readable code list information and its interchange between systems”[^1]. Genericode is developed by the [Organization for the Advancement of Structured Information Standards (OASIS)](https://www.oasis-open.org/). See the [website of the OASIS Code List Representation TC](https://www.oasis-open.org/committees/codelist/) for more information.

[^1]: Source: [Genericode Approved as an OASIS Standard](https://www.oasis-open.org/2023/02/01/genericode-approved-as-an-oasis-standard/)

### OpenDocument format

The [OpenDocument Format](https://docs.oasis-open.org/office/OpenDocument/v1.3/OpenDocument-v1.3-part1-introduction.html) is “a free, open XML-based document file format for office applications, to be used for documents containing text, spreadsheets, charts, and graphical elements”[^2]. The OpenDocument Format is also developed by OASIS. See the [website of the OASIS Open Document Format for Office Applications (OpenDocument) TC](https://www.oasis-open.org/committees/office/) and the [Wikipedia article on OpenDocument](https://en.wikipedia.org/wiki/OpenDocument) for more information.

genericode-xml-filter uses version 1.3 of the OpenDocument specification.

[^2]: Source: [OpenDocument V1.3 OASIS Standard published](https://www.oasis-open.org/2021/06/16/opendocument-v1-3-oasis-standard-published/)

### LibreOffice

[LibreOffice](https://www.libreoffice.org/) is an office suite that uses OpenDocument files as its native file format[^3]. Through the use of [XML filters](https://help.libreoffice.org/latest/en-US/text/shared/guide/xsltfilter.html), files in other file formats can be loaded and saved transparently.

[^3]: Source: [XML File Formats](https://help.libreoffice.org/latest/en-US/text/shared/00/00000021.html)

## Installation

Prerequisite: LibreOffice is installed, the minimum version needed is 24.2.5[^4].

[^4]: This is because of a [bug in the XML filter functionality](https://bugs.documentfoundation.org/show_bug.cgi?id=161789) in earlier versions.

1. Download the [latest release of the genericode XML filter](https://github.com/SDFIdk/genericode-xml-filter/releases/latest)
2. Open LibreOffice Calc
3. Go to Tools -> XML Filter Settings... (for more information about this dialog, see the [LibreOffice Help page on XML Filter Settings](https://help.libreoffice.org/latest/en-US/text/shared/01/06150000.html))
4. Choose Open Package...
5. Open the jar file downloaded in step 1

## Usage

### Opening genericode files

#### Via the LibreOffice Calc GUI

1. Access the Open File dialog (for more information about this dialog, see the [LibreOffice Help page on opening files](https://help.libreoffice.org/latest/en-US/text/shared/01/01020000.html))
2. Select the “Code List Representation (genericode)” file type
3. Choose the genericode file you want to open

> [!IMPORTANT]
> The “Code List Representation (genericode)” file type must be chosen instead of “All files (*.*)”, otherwise the file will not be opened correctly.

#### Via the command line prompt

```bat
"C:\Program Files\LibreOffice\program\scalc.exe" codelist.gc --infilter="Genericode"
```

### Editing genericode files

The spreadsheet contains four sheets:

1. Identification: displays the contents of `/gc:CodeList/Identification`
2. Additional code list metadata: displays the contents of `/gc:CodeList/Annotation/Description`
3. Columns: displays the contents of `/gc:CodeList/ColumnSet`
4. Values: displays the contents of `/gc:CodeList/SimpleCodeList`

### Profile of genericode

Not all possibilities in the genericode standard are supported, and selected additional metadata element on the code list and its columns are supported.

Conventions and limitations:

- For sheets Identification and Additional code list metadata: the first column contains the element name and any attributes in the format `ElementName att1=value1,att2=value2`, e.g. `AlternateFormatLocationUri MimeType=text/csv` or `dcterms:description xml:lang`. Make sure not to have a space after the element name if the element should not have any attributes.
- The Qualified DC XML Schemas are used for additional metadata, see also https://www.dublincore.org/schemas/xmls/, the properties from the `http://purl.org/dc/terms/` namespace are used.
- For the code list itself, all `dcterms` properties are supported in `/gc:CodeList/Annotation/Description`, and `xml:lang` is  supported.
- The `Id` attribute and the `ShortName` of a column are equal. If a column has another value as `ShortName` than the `Id` attribute in the genericode file, it will be replaced with the value of the `Id` attribute when the genericode file is saved.
- For columns, only `@Id`, `@Use`, `Data/@Type`, `Data/@Lang` and `Annotation/Description/dcterms:description` are supported. `ShortName` is supported as well but is always set to the same value as the `Id` attribute and therefore not visible in the spreadsheet. `xml:lang` is not supported.
- The `Id` attributes of the columns are used for the column headers in sheet Values.
- An undefined value is encoded as an empty `Value` element.
- The order of the `Value` elements must be the same as the order of the `Column` elements, if not, the transformation will be terminated and LibreOffice will display an error.

An example:

```xml
<gc:CodeList xmlns:gc="http://docs.oasis-open.org/codelist/ns/genericode/1.0/" xmlns:dcterms="http://purl.org/dc/terms/">
  <Annotation>
    <Description>
      <dcterms:description>Description of my code list.</dcterms:description>
      <dcterms:language>en</dcterms:language>
      <dcterms:license>https://path/to/license</dcterms:license>
    </Description>
  </Annotation>
  <Identification>
    <ShortName>mytype</ShortName>
    <Version>1.0.0</Version>
    <CanonicalUri>urn:uuid:08b6a53e-3d1b-42da-a6fb-4741a9ba5828</CanonicalUri>
    <CanonicalVersionUri>urn:uuid:ceb73150-b5c4-4d61-8b18-ccdb02e1acde</CanonicalVersionUri>
    <LocationUri>https://path/to/v1.0.0.mytype.gc</LocationUri>
    <AlternateFormatLocationUri MimeType="text/csv">https://path/to/v1.0.0.mytype.csv</AlternateFormatLocationUri>
    <AlternateFormatLocationUri MimeType="text/html">https://path/to/v1.0.0.mytype.html</AlternateFormatLocationUri>
    <Agency>
      <LongName>My organisation</LongName>
    </Agency>
  </Identification>
  <ColumnSet>
    <Column Id="code" Use="required">
      <Annotation>
        <Description>
          <dcterms:description>allowed value in data</dcterms:description>
        </Description>
      </Annotation>
      <ShortName>code</ShortName>
      <Data Type="string"/>
    </Column>
    <Column Id="name" Use="required">
      <Annotation>
        <Description>
          <dcterms:description>word or words that the thing is known by</dcterms:description>
        </Description>
      </Annotation>
      <ShortName>name</ShortName>
      <Data Type="string" Lang="en"/>
    </Column>
    <Key Id="key_code">
      <ShortName>key_code</ShortName>
      <ColumnRef Ref="code"/>
    </Key>
    <Key Id="key_name">
      <ShortName>key_name</ShortName>
      <ColumnRef Ref="name"/>
    </Key>
  </ColumnSet>
  <SimpleCodeList/>
</gc:CodeList>

```

## Development

Further development of the XML filter requires the set-up of the XML filter in LibreOffice Calc via Tools -> XML Filter Settings..., see also [Creating XML Filters](https://help.libreoffice.org/latest/en-US/text/shared/guide/xsltfilter_create.html).

General tab:

- Filter name: Genericode
- Application: LibreOffice Calc (.ods)
- Name of file type: Code List Representation (genericode)
- File extension: gc

Transformation tab:

- DocType: (leave empty)
- XSLT for export: location of [ods2gc.xsl](src/ods2gc.xsl) in your local working copy
- XSLT for import: location of [gc2ods.xsl](src/gc2ods.xsl) in your local working copy
- Template for import: (leave empty)

LibreOffice uses the [libxslt](https://gitlab.gnome.org/GNOME/libxslt/-/wikis/home) library. This can be verified by adding the following comment to the output somewhere. libxslt provides an implementation of [XSLT 1.0](https://www.w3.org/TR/xslt-10/) and also supports most [EXSLT extension functions](https://exslt.github.io/).

```xml
<xsl:comment>
    <xsl:value-of select="system-property('xsl:vendor')" />
</xsl:comment>
```

Therefore, the XSLT stylesheets can be tested using [xsltproc](https://gnome.pages.gitlab.gnome.org/libxslt/xsltproc.html), the command line tool for libxslt.

For testing [ods2gc.xsl](src/ods2gc.xsl), save the code list you are testing with as a Flat XML ODF Spreadsheet (.fods file extension). Investigate that file using an XML editor and use it as input to the XSLT transformation.

```bat
xsltproc -o path\to\output.gc path\to\ods2gc.xsl path\to\codelist.fods
```

For testing [gc2ods.xsl](src/gc2ods.xsl), transform the genericode code list to a Flat XML ODF Spreadsheet (.fods file extension). Validate the spreadsheet against the [OpenDocument RELAX NG schema](https://docs.oasis-open.org/office/OpenDocument/v1.3/os/schemas/OpenDocument-v1.3-schema.rng), e.g. with [xmllint](https://gnome.pages.gitlab.gnome.org/libxml2/xmllint.html), and verify that it looks as it should in LibreOffice.

```bat
xsltproc -o path\to\output.fods path\to\gc2ods.xsl path\to\codelist.gc
```

```bat
xmllint --noout --relaxng path\to\OpenDocument-v1.3-schema.rng path\to\output.fods
```

> [!TIP]
> xsltproc and xmllint can be installed as a package of [OSGeo4W](https://trac.osgeo.org/osgeo4w/).

For more information about creating XML filters, see:

- EISENBERG, J. David. _[OASIS OpenDocument Essentials: Using OASIS OpenDocument XML](https://archive.org/details/ODEssentials)_. 2005. ISBN ISBN 1-4116-6832-4.
- [LibreOffice SDK Guide: Chapter 50. Importing XML](https://wiki.documentfoundation.org/Documentation/SDKGuide/Importing_XML)
- [Create XSLT filters for import and export](https://forum.openoffice.org/en/forum/viewtopic.php?t=3490)

## Building

Save the XML filter as a package (a jar file) following the instructions on [Distributing An XML Filter As Package](https://help.libreoffice.org/latest/en-US/text/shared/guide/xsltfilter_distribute.html).