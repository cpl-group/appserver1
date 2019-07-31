<?xml version="1.0"?> 
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0" xmlns:genergy="/">
	<xsl:output method="html" indent="yes"/>
	<xsl:template match="genergymenu">
		<br/>
		<genergy>
		<genergy:root id="root">
<!-- 			<xsl:attribute name="node">true</xsl:attribute>
			<xsl:attribute name="imgExpand">/images/folderclosed.gif</xsl:attribute>
			<xsl:attribute name="imgClose">/images/folderclosed.gif</xsl:attribute>
			<xsl:value-of select="@Corp_name"/> -->
			<xsl:apply-templates/>
		</genergy:root>
		</genergy>
	</xsl:template>
	
	<xsl:template match="branch">
		<genergy:branch>
			<xsl:attribute name="node">true</xsl:attribute>
			<xsl:attribute name="name">node</xsl:attribute>
			<xsl:attribute name="nid"><xsl:value-of select="@nid"/></xsl:attribute>
			<xsl:attribute name="fid"><xsl:value-of select="@fid"/></xsl:attribute>
			<xsl:attribute name="imgExpand">/images/folderclosed.gif</xsl:attribute>
			<xsl:attribute name="imgClose">/images/folderclosed.gif</xsl:attribute>
			<a><xsl:attribute name="name"><xsl:value-of select="@nid"/></xsl:attribute>
				<xsl:if test="@link or @onclick">
				<xsl:attribute name="onclick"><xsl:value-of select="@onclick"/></xsl:attribute>
				<xsl:attribute name="target"><xsl:value-of select="@target"/></xsl:attribute>
				<xsl:attribute name="href"><xsl:value-of select="@link"/></xsl:attribute>
			</xsl:if><xsl:value-of select="@label"/></a><br/><null></null>
			<xsl:apply-templates/>
		</genergy:branch>
	</xsl:template>

	<xsl:template match="leaf">
		<genergy:leaf>
			<xsl:attribute name="node">true</xsl:attribute>
			<xsl:attribute name="name">node</xsl:attribute>
			<xsl:attribute name="nid"><xsl:value-of select="@nid"/></xsl:attribute>
			<xsl:attribute name="fid"><xsl:value-of select="@fid"/></xsl:attribute>
			<xsl:attribute name="imgExpand">/images/msie_doc.gif</xsl:attribute>
			<xsl:attribute name="imgClose">/images/msie_doc.gif</xsl:attribute>
			<a><xsl:attribute name="name"><xsl:value-of select="@nid"/></xsl:attribute>
				<xsl:if test="@link or @onclick">
				<xsl:attribute name="onclick"><xsl:value-of select="@onclick"/></xsl:attribute>
				<xsl:attribute name="target"><xsl:value-of select="@target"/></xsl:attribute>
				<xsl:attribute name="href"><xsl:value-of select="@link"/></xsl:attribute>
			</xsl:if><xsl:value-of select="@label"/></a><br/><null></null>
		</genergy:leaf>
	</xsl:template>
	
</xsl:stylesheet>