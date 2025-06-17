use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
#[derive(Default)]
pub enum SlideTemplate {
    #[default]
    Default,
    Professional,
    Modern,
    Minimal,
    Custom(CustomTemplate),
}


#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct CustomTemplate {
    pub name: String,
    pub theme_colors: ThemeColors,
    pub fonts: FontScheme,
    pub layout_settings: LayoutSettings,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ThemeColors {
    pub background: String,
    pub text_primary: String,
    pub text_secondary: String,
    pub accent_1: String,
    pub accent_2: String,
    pub accent_3: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct FontScheme {
    pub title_font: String,
    pub body_font: String,
    pub code_font: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct LayoutSettings {
    pub slide_width: i32,
    pub slide_height: i32,
    pub margin_top: i32,
    pub margin_bottom: i32,
    pub margin_left: i32,
    pub margin_right: i32,
    pub title_height: i32,
    pub content_spacing: i32,
}

impl SlideTemplate {
    pub fn from_name(name: &str) -> Self {
        match name.to_lowercase().as_str() {
            "default" => SlideTemplate::Default,
            "professional" => SlideTemplate::Professional,
            "modern" => SlideTemplate::Modern,
            "minimal" => SlideTemplate::Minimal,
            _ => SlideTemplate::Default,
        }
    }

    pub fn get_theme_colors(&self) -> ThemeColors {
        match self {
            SlideTemplate::Default => ThemeColors {
                background: "FFFFFF".to_string(),
                text_primary: "000000".to_string(),
                text_secondary: "666666".to_string(),
                accent_1: "4F81BD".to_string(),
                accent_2: "F79646".to_string(),
                accent_3: "9BBB59".to_string(),
            },
            SlideTemplate::Professional => ThemeColors {
                background: "FFFFFF".to_string(),
                text_primary: "1F1F1F".to_string(),
                text_secondary: "757575".to_string(),
                accent_1: "2E75B6".to_string(),
                accent_2: "C65911".to_string(),
                accent_3: "70AD47".to_string(),
            },
            SlideTemplate::Modern => ThemeColors {
                background: "F8F9FA".to_string(),
                text_primary: "212529".to_string(),
                text_secondary: "6C757D".to_string(),
                accent_1: "007BFF".to_string(),
                accent_2: "FD7E14".to_string(),
                accent_3: "28A745".to_string(),
            },
            SlideTemplate::Minimal => ThemeColors {
                background: "FFFFFF".to_string(),
                text_primary: "2C2C2C".to_string(),
                text_secondary: "8C8C8C".to_string(),
                accent_1: "007ACC".to_string(),
                accent_2: "FF6B35".to_string(),
                accent_3: "32CD32".to_string(),
            },
            SlideTemplate::Custom(template) => template.theme_colors.clone(),
        }
    }

    pub fn get_fonts(&self) -> FontScheme {
        match self {
            SlideTemplate::Default => FontScheme {
                title_font: "Calibri".to_string(),
                body_font: "Calibri".to_string(),
                code_font: "Consolas".to_string(),
            },
            SlideTemplate::Professional => FontScheme {
                title_font: "Segoe UI".to_string(),
                body_font: "Segoe UI".to_string(),
                code_font: "Consolas".to_string(),
            },
            SlideTemplate::Modern => FontScheme {
                title_font: "Roboto".to_string(),
                body_font: "Roboto".to_string(),
                code_font: "Fira Code".to_string(),
            },
            SlideTemplate::Minimal => FontScheme {
                title_font: "Helvetica".to_string(),
                body_font: "Helvetica".to_string(),
                code_font: "Monaco".to_string(),
            },
            SlideTemplate::Custom(template) => template.fonts.clone(),
        }
    }

    pub fn get_layout_settings(&self) -> LayoutSettings {
        let base_settings = LayoutSettings {
            slide_width: 9144000,  // 10 inches in EMUs
            slide_height: 6858000, // 7.5 inches in EMUs
            margin_top: 457200,    // 0.5 inches
            margin_bottom: 457200,
            margin_left: 685800, // 0.75 inches
            margin_right: 685800,
            title_height: 1143000,   // 1.25 inches
            content_spacing: 228600, // 0.25 inches
        };

        match self {
            SlideTemplate::Minimal => LayoutSettings {
                margin_top: 914400, // 1 inch for more space
                margin_bottom: 914400,
                margin_left: 1371600, // 1.5 inches for more space
                margin_right: 1371600,
                ..base_settings
            },
            SlideTemplate::Custom(template) => template.layout_settings.clone(),
            _ => base_settings,
        }
    }

    pub fn get_slide_master_xml(&self) -> String {
        let colors = self.get_theme_colors();
        let fonts = self.get_fonts();

        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
    <p:cSld>
        <p:bg>
            <p:bgRef idx="1001">
                <a:srgbClr val="{}"/>
            </p:bgRef>
        </p:bg>
        <p:spTree>
            <p:nvGrpSpPr>
                <p:cNvPr id="1" name=""/>
                <p:cNvGrpSpPr/>
                <p:nvPr/>
            </p:nvGrpSpPr>
            <p:grpSpPr>
                <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="0" cy="0"/>
                    <a:chOff x="0" y="0"/>
                    <a:chExt cx="0" cy="0"/>
                </a:xfrm>
            </p:grpSpPr>
        </p:spTree>
    </p:cSld>
    <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
    <p:sldLayoutIdLst>
        <p:sldLayoutId id="2147483649" r:id="rId1"/>
    </p:sldLayoutIdLst>
    <p:txStyles>
        <p:titleStyle>
            <a:lvl1pPr>
                <a:defRPr sz="4400" kern="1200">
                    <a:solidFill>
                        <a:srgbClr val="{}"/>
                    </a:solidFill>
                    <a:latin typeface="{}"/>
                </a:defRPr>
            </a:lvl1pPr>
        </p:titleStyle>
        <p:bodyStyle>
            <a:lvl1pPr>
                <a:defRPr sz="2800">
                    <a:solidFill>
                        <a:srgbClr val="{}"/>
                    </a:solidFill>
                    <a:latin typeface="{}"/>
                </a:defRPr>
            </a:lvl1pPr>
        </p:bodyStyle>
        <p:otherStyle>
            <a:lvl1pPr>
                <a:defRPr>
                    <a:solidFill>
                        <a:srgbClr val="{}"/>
                    </a:solidFill>
                    <a:latin typeface="{}"/>
                </a:defRPr>
            </a:lvl1pPr>
        </p:otherStyle>
    </p:txStyles>
</p:sldMaster>"#,
            colors.background,
            colors.text_primary,
            fonts.title_font,
            colors.text_primary,
            fonts.body_font,
            colors.text_secondary,
            fonts.body_font
        )
    }

    pub fn get_theme_xml(&self) -> String {
        let colors = self.get_theme_colors();
        let fonts = self.get_fonts();

        format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom Theme">
    <a:themeElements>
        <a:clrScheme name="Custom">
            <a:dk1>
                <a:srgbClr val="{}"/>
            </a:dk1>
            <a:lt1>
                <a:srgbClr val="{}"/>
            </a:lt1>
            <a:dk2>
                <a:srgbClr val="{}"/>
            </a:dk2>
            <a:lt2>
                <a:srgbClr val="EEECE1"/>
            </a:lt2>
            <a:accent1>
                <a:srgbClr val="{}"/>
            </a:accent1>
            <a:accent2>
                <a:srgbClr val="{}"/>
            </a:accent2>
            <a:accent3>
                <a:srgbClr val="{}"/>
            </a:accent3>
            <a:accent4>
                <a:srgbClr val="8064A2"/>
            </a:accent4>
            <a:accent5>
                <a:srgbClr val="4BACC6"/>
            </a:accent5>
            <a:accent6>
                <a:srgbClr val="F39646"/>
            </a:accent6>
            <a:hlink>
                <a:srgbClr val="0000FF"/>
            </a:hlink>
            <a:folHlink>
                <a:srgbClr val="800080"/>
            </a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Custom">
            <a:majorFont>
                <a:latin typeface="{}"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
            </a:majorFont>
            <a:minorFont>
                <a:latin typeface="{}"/>
                <a:ea typeface=""/>
                <a:cs typeface=""/>
            </a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Office">
            <a:fillStyleLst>
                <a:solidFill>
                    <a:schemeClr val="phClr"/>
                </a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="50000"/>
                                <a:satMod val="300000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="35000">
                            <a:schemeClr val="phClr">
                                <a:tint val="37000"/>
                                <a:satMod val="300000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:tint val="15000"/>
                                <a:satMod val="350000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="16200000" scaled="1"/>
                </a:gradFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:shade val="51000"/>
                                <a:satMod val="130000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="80000">
                            <a:schemeClr val="phClr">
                                <a:shade val="93000"/>
                                <a:satMod val="130000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:shade val="94000"/>
                                <a:satMod val="135000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:lin ang="16200000" scaled="0"/>
                </a:gradFill>
            </a:fillStyleLst>
            <a:lnStyleLst>
                <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr">
                            <a:shade val="95000"/>
                            <a:satMod val="105000"/>
                        </a:schemeClr>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                </a:ln>
                <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                </a:ln>
                <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
                    <a:solidFill>
                        <a:schemeClr val="phClr"/>
                    </a:solidFill>
                    <a:prstDash val="solid"/>
                </a:ln>
            </a:lnStyleLst>
            <a:effectStyleLst>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
                            <a:srgbClr val="000000">
                                <a:alpha val="38000"/>
                            </a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
                            <a:srgbClr val="000000">
                                <a:alpha val="35000"/>
                            </a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
                <a:effectStyle>
                    <a:effectLst>
                        <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
                            <a:srgbClr val="000000">
                                <a:alpha val="35000"/>
                            </a:srgbClr>
                        </a:outerShdw>
                    </a:effectLst>
                </a:effectStyle>
            </a:effectStyleLst>
            <a:bgFillStyleLst>
                <a:solidFill>
                    <a:schemeClr val="phClr"/>
                </a:solidFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="40000"/>
                                <a:satMod val="350000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="40000">
                            <a:schemeClr val="phClr">
                                <a:tint val="45000"/>
                                <a:shade val="99000"/>
                                <a:satMod val="350000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:shade val="20000"/>
                                <a:satMod val="255000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:path path="circle">
                        <a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
                    </a:path>
                </a:gradFill>
                <a:gradFill rotWithShape="1">
                    <a:gsLst>
                        <a:gs pos="0">
                            <a:schemeClr val="phClr">
                                <a:tint val="80000"/>
                                <a:satMod val="300000"/>
                            </a:schemeClr>
                        </a:gs>
                        <a:gs pos="100000">
                            <a:schemeClr val="phClr">
                                <a:shade val="30000"/>
                                <a:satMod val="200000"/>
                            </a:schemeClr>
                        </a:gs>
                    </a:gsLst>
                    <a:path path="circle">
                        <a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
                    </a:path>
                </a:gradFill>
            </a:bgFillStyleLst>
        </a:fmtScheme>
    </a:themeElements>
    <a:objectDefaults/>
    <a:extraClrSchemeLst/>
</a:theme>"#,
            colors.text_primary,
            colors.background,
            colors.text_secondary,
            colors.accent_1,
            colors.accent_2,
            colors.accent_3,
            fonts.title_font,
            fonts.body_font
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_template_from_name() {
        assert!(matches!(
            SlideTemplate::from_name("default"),
            SlideTemplate::Default
        ));
        assert!(matches!(
            SlideTemplate::from_name("professional"),
            SlideTemplate::Professional
        ));
        assert!(matches!(
            SlideTemplate::from_name("modern"),
            SlideTemplate::Modern
        ));
        assert!(matches!(
            SlideTemplate::from_name("minimal"),
            SlideTemplate::Minimal
        ));
        assert!(matches!(
            SlideTemplate::from_name("unknown"),
            SlideTemplate::Default
        ));
    }

    #[test]
    fn test_theme_colors() {
        let template = SlideTemplate::Professional;
        let colors = template.get_theme_colors();
        assert_eq!(colors.text_primary, "1F1F1F");
        assert_eq!(colors.accent_1, "2E75B6");
    }

    #[test]
    fn test_fonts() {
        let template = SlideTemplate::Modern;
        let fonts = template.get_fonts();
        assert_eq!(fonts.title_font, "Roboto");
        assert_eq!(fonts.code_font, "Fira Code");
    }
}
