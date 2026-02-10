#ifndef RESOURCE_H
#define RESOURCE_H

#ifndef IDC_STATIC
#define IDC_STATIC -1
#endif

// Main Resources
#define IDI_ICON1                       101
#define IDR_WAV_LIGHTS_ON               102
#define IDR_WAV_LIGHTS_OFF              103
#define IDR_MSI_DLL                     104

// Dialogs
#define IDD_SETTINGS                    105
#define IDD_ABOUT                       106

// Tray Menu
#define ID_TRAY_EXIT                    1001
#define ID_TRAY_CONFIG                  2001
#define ID_TRAY_LOG                     3001
#define ID_TRAY_ABOUT                   4001
#define ID_TRAY_TOGGLE_LEDS				4002

// --- SETTINGS DIALOG CONTROLS ---

// Tabs & Main Buttons
#define IDC_TAB_PROFILES                2000
#define IDC_CHK_STARTUP                 2001
#define IDC_BTN_RESET                   2002
#define IDC_BTN_RESET_ADVANCED          2003

// Profile Inputs
#define IDC_EDIT_LABEL                  2100
#define IDC_CHK_ACTIVE_PROFILE          2101
#define IDC_COMBO_DEVICE                2102
#define IDC_COMBO_AREA                  2103
#define IDC_SENSOR_ID                   2104
#define IDC_EDIT_SERVER                 2105
#define IDC_TEMP_LOW                    2106
#define IDC_TEMP_MED                    2107
#define IDC_TEMP_HIGH                   2108
#define IDC_HEX_LOW                     2109
#define IDC_HEX_MED                     2110
#define IDC_HEX_HIGH                    2111
#define IDC_COMBO_SENSOR_UPDATE         2112
#define IDC_COMBO_LED_FPS               2113
#define IDC_COMBO_SMOOTHING             2114

// Profile Groups (CRITICAL: Added IDC_GRP_THRESHOLDS)
#define IDC_GRP_MSI                     2200
#define IDC_GRP_SOURCE                  2201
#define IDC_GRP_ADVANCED                2202
#define IDC_GRP_THRESHOLDS              2203 

// Profile Labels (Static text needs IDs to be hidden)
#define IDC_LBL_PROFILE_NAME            2300
#define IDC_LBL_DEVICE                  2301
#define IDC_LBL_AREA                    2302
#define IDC_LBL_TEMP_SENS               2303
#define IDC_LBL_WEB_SRV                 2304
#define IDC_LBL_LOW                     2305
#define IDC_LBL_MED                     2306
#define IDC_LBL_HIGH                    2307
#define IDC_LBL_C1                      2308
#define IDC_LBL_C2                      2309
#define IDC_LBL_C3                      2310
#define IDC_LBL_HEX1                    2311
#define IDC_LBL_HEX2                    2312
#define IDC_LBL_HEX3                    2313
#define IDC_LBL_ADV_SENSOR              2314
#define IDC_LBL_ADV_LED                 2315
#define IDC_LBL_ADV_SMOOTH              2316

// --- SHORTCUTS LAYER ---
#define IDC_GRP_SHORTCUTS               3000
#define IDC_LBL_TOGGLE                  3001
#define IDC_HK_TOGGLE                   3002
#define IDC_LBL_P1                      3003
#define IDC_HK_P1                       3004
#define IDC_LBL_P2                      3005
#define IDC_HK_P2                       3006
#define IDC_LBL_P3                      3007
#define IDC_HK_P3                       3008
#define IDC_LBL_P4                      3009
#define IDC_HK_P4                       3010
#define IDC_LBL_P5                      3011
#define IDC_HK_P5                       3012
#define IDC_BTN_RESET_SHORTCUTS         3013

// About Dialog
#define IDC_ABOUT_ICON                  4001
#define IDC_ABOUT_VERSION               4002
#define IDC_GITHUB_LINK                 4003

#endif // RESOURCE_H
