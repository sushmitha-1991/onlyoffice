<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0"
                xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml">

    <xsl:output method="html" indent="yes" encoding="UTF-8"/>

    <xsl:template match="/">

        <html>
            <head>
                <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
            </head>
            <body>
                <xsl:apply-templates select="//w:body"/>
            </body>
        </html>

    </xsl:template>

    <!-- Handle paragraph -->
    <xsl:template match="w:p">

        <p>
            <!-- Alignment -->
            <xsl:if test="w:pPr/w:jc">
                <xsl:attribute name="style">
                    text-align:
                    <xsl:choose>
                        <xsl:when test="w:pPr/w:jc/@w:val = 'center'">center</xsl:when>
                        <xsl:when test="w:pPr/w:jc/@w:val = 'right'">right</xsl:when>
                        <xsl:when test="w:pPr/w:jc/@w:val = 'both'">justify</xsl:when>
                        <xsl:otherwise>left</xsl:otherwise>
                    </xsl:choose>;
                </xsl:attribute>
            </xsl:if>

            <xsl:apply-templates/>

        </p>

    </xsl:template>

    <!-- Handle run -->
    <xsl:template match="w:r">

        <span>
            <xsl:attribute name="style">

                <!-- Font Family -->
                <xsl:if test="w:rPr/w:rFonts/@w:ascii">
                    font-family: <xsl:value-of select="w:rPr/w:rFonts/@w:ascii"/>;
                </xsl:if>

                <!-- Font Size -->
                <xsl:if test="w:rPr/w:sz/@w:val">
                    font-size: <xsl:value-of select="w:rPr/w:sz/@w:val div 2"/>pt;
                </xsl:if>

                <!-- Bold -->
                <xsl:choose>
                    <xsl:when test="w:rPr/w:b">
                        font-weight: bold;
                    </xsl:when>
                    <xsl:otherwise>
                        font-weight: normal;
                    </xsl:otherwise>
                </xsl:choose>

                <!-- Italic -->
                <xsl:choose>
                    <xsl:when test="w:rPr/w:i">
                        font-style: italic;
                    </xsl:when>
                    <xsl:otherwise>
                        font-style: normal;
                    </xsl:otherwise>
                </xsl:choose>

            </xsl:attribute>

            <xsl:apply-templates/>

        </span>

    </xsl:template>
    <xsl:template match="w:p[contains(normalize-space(string(.)), 'Uncorrected/Not for publication')]" priority="10"/>
    <!-- Handle text -->
    <xsl:template match="w:t">
        <xsl:value-of select="."/>
    </xsl:template>

</xsl:stylesheet>
