#define _WIN32_WINNT 0x0501
#define _WIN32_IE 0x0500
#include <windows.h>
#include <commctrl.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <commdlg.h>

#define MAX_LINE 4096
#define MAX_ITEMS 10000
#define ID_LEN 256
#define SKU_LEN 256

// --- Omnichannel Config Structure ---
typedef struct {
    char name[128];
    char map_file[256];
    char id_col[128];
    char qty_col[128];
    char webhook[2048];
} PlatformConfig;

PlatformConfig platforms[20];
int platform_count = 0;
int current_platform_idx = 0;
char config_path[MAX_PATH];

// --- Data Structures ---
typedef struct {
    char key[ID_LEN];
    char sku[SKU_LEN];
} Mapping;

typedef struct {
    char sku[SKU_LEN];
    int deduction;
} Deduction;

Mapping mappings[MAX_ITEMS];
int mapping_count = 0;
Deduction deductions[MAX_ITEMS];
int deduction_count = 0;

char mapping_file_path[MAX_PATH] = "";
char sales_file_path[MAX_PATH] = "";
char google_webhook_url[2048] = "https://script.google.com/macros/s/AKfycbzMcc-tgFb1BGr_qDXBk0FqYJemYx7BDyijWXH-ZrS0o3hcqwqH7dEJm1AT_Wg3D1PDtQ/exec";

// Global GUI
HWND hMainWnd, hStatus, hMapPath, hSalesPath, hComboPlatform;
HWND hBtnProcess, hBtnSync;
HWND hEditorWnd = NULL, hEditorListView, hEditCode;
HWND hResultListView;

// CSV memory map
char* csv_headers[50];
int csv_col_count = 0;
char* csv_data[MAX_ITEMS][50];
int csv_row_count = 0;
int selected_row = -1;

// --- Helper Functions ---
wchar_t* utf8_to_wide(const char* utf8) {
    if (!utf8) return NULL;
    int wlen = MultiByteToWideChar(CP_UTF8, 0, utf8, -1, NULL, 0);
    if (wlen <= 0) return NULL;
    wchar_t* wbuf = malloc(wlen * sizeof(wchar_t));
    MultiByteToWideChar(CP_UTF8, 0, utf8, -1, wbuf, wlen);
    return wbuf;
}

char* wide_to_utf8(const wchar_t* wstr) {
    if (!wstr) return NULL;
    int ulen = WideCharToMultiByte(CP_UTF8, 0, wstr, -1, NULL, 0, NULL, NULL);
    if (ulen <= 0) return NULL;
    char* ubuf = malloc(ulen);
    WideCharToMultiByte(CP_UTF8, 0, wstr, -1, ubuf, ulen, NULL, NULL);
    return ubuf;
}

void add_log(const char* text) {
    if (!hStatus) return;
    int length = GetWindowTextLength(hStatus);
    SendMessage(hStatus, EM_SETSEL, (WPARAM)length, (LPARAM)length);
    wchar_t* wtxt = utf8_to_wide(text);
    if (wtxt) {
        SendMessageW(hStatus, EM_REPLACESEL, 0, (LPARAM)wtxt);
        free(wtxt);
    }
    SendMessageW(hStatus, EM_REPLACESEL, 0, (LPARAM)L"\r\n");
}

void trim(char* str) {
    if (!str) return;
    char *p = str;
    int l = strlen(p);
    while(l > 0 && (p[l - 1] == '\r' || p[l - 1] == '\n' || p[l - 1] == ' ' || p[l - 1] == '\"')) p[--l] = 0;
    while(*p && (*p == ' ' || *p == '\"')) p++;
    memmove(str, p, l + 1);
}

char* next_field(char** cursor) {
    if (!cursor || !*cursor || **cursor == '\0' || **cursor == '\r' || **cursor == '\n') return NULL;
    char* start = *cursor;
    if (*start == '\"') {
        start++; char* p = start; char* out = start;
        while (*p) {
            if (*p == '\"') {
                if (*(p + 1) == '\"') { *out++ = '\"'; p += 2; } 
                else {
                    *out = '\0'; p++;
                    while (*p && *p != ',' && *p != '\r' && *p != '\n') p++;
                    if (*p == ',') p++; 
                    *cursor = p; return start;
                }
            } else *out++ = *p++;
        }
        *out = '\0'; *cursor = p; return start;
    } else {
        char* end = strchr(start, ',');
        if (end) { *end = '\0'; *cursor = end + 1; } 
        else *cursor = start + strlen(start);
        return start;
    }
}

const char* find_sku(const char* key) {
    for (int i = 0; i < mapping_count; i++) {
        if (strcmp(mappings[i].key, key) == 0) return mappings[i].sku;
    }
    return NULL;
}

// --- Config Parser ---
void load_config() {
    GetCurrentDirectoryA(MAX_PATH, config_path);
    strcat(config_path, "\\config.ini");
    char buf[128];
    GetPrivateProfileStringA("Platforms", "Count", "0", buf, sizeof(buf), config_path);
    platform_count = atoi(buf);
    if (platform_count == 0) {
        MessageBox(NULL, "Warning: config.ini is missing or invalid. Please configure platforms.", "Configuration Error", MB_ICONWARNING);
    }
    for (int i = 0; i < platform_count; i++) {
        char key[16]; sprintf(key, "%d", i + 1);
        GetPrivateProfileStringA("Platforms", key, "Unknown", platforms[i].name, sizeof(platforms[i].name), config_path);
        
        GetPrivateProfileStringA(platforms[i].name, "MappingFile", "mapping.csv", platforms[i].map_file, sizeof(platforms[i].map_file), config_path);
        GetPrivateProfileStringA(platforms[i].name, "SalesIdColumnHeader", "SKU", platforms[i].id_col, sizeof(platforms[i].id_col), config_path);
        GetPrivateProfileStringA(platforms[i].name, "SalesQtyColumnHeader", "Quantity", platforms[i].qty_col, sizeof(platforms[i].qty_col), config_path);
        GetPrivateProfileStringA(platforms[i].name, "GoogleWebhookUrl", "", platforms[i].webhook, sizeof(platforms[i].webhook), config_path);
    }
}

// --- Editor Functions ---
void load_csv_memory() {
    for(int r=0; r<csv_row_count; r++) for(int c=0; c<csv_col_count; c++) { free(csv_data[r][c]); csv_data[r][c] = NULL; }
    for(int c=0; c<csv_col_count; c++) { free(csv_headers[c]); csv_headers[c] = NULL; }
    csv_row_count = 0; csv_col_count = 0;
    
    FILE* f = fopen(mapping_file_path, "r");
    if(!f) return;
    char line[MAX_LINE];
    if (fgets(line, MAX_LINE, f)) {
        char* header = line;
        if ((unsigned char)header[0] == 0xEF && (unsigned char)header[1] == 0xBB && (unsigned char)header[2] == 0xBF) header += 3;
        char buf[MAX_LINE]; strcpy(buf, header);
        char* cursor = buf;
        while(csv_col_count < 50) {
            char* val = next_field(&cursor);
            if(!val) break;
            trim(val); csv_headers[csv_col_count++] = strdup(val);
        }
    }
    while (fgets(line, MAX_LINE, f) && csv_row_count < MAX_ITEMS) {
        char buf[MAX_LINE]; strcpy(buf, line);
        char* cursor = buf;
        for(int c=0; c<csv_col_count; c++) {
            char* val = next_field(&cursor);
            if(val) { trim(val); csv_data[csv_row_count][c] = strdup(val); }
            else csv_data[csv_row_count][c] = strdup("");
        }
        csv_row_count++;
    }
    fclose(f);
}

void populate_editor_listview() {
    ListView_DeleteAllItems(hEditorListView);
    HWND hHeader = ListView_GetHeader(hEditorListView);
    int cols = Header_GetItemCount(hHeader);
    for(int i = cols - 1; i >= 0; i--) ListView_DeleteColumn(hEditorListView, i);
    
    LVCOLUMNW lvcw; lvcw.mask = LVCF_FMT | LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
    for(int c=0; c<csv_col_count; c++) {
        wchar_t* wtxt = utf8_to_wide(csv_headers[c]);
        lvcw.iSubItem = c; lvcw.pszText = wtxt ? wtxt : L"?"; 
        lvcw.cx = (c == 0 || c == csv_col_count-1) ? 200 : 100; lvcw.fmt = LVCFMT_LEFT;
        SendMessageW(hEditorListView, LVM_INSERTCOLUMNW, c, (LPARAM)&lvcw);
        if (wtxt) free(wtxt);
    }
    for(int r=0; r<csv_row_count; r++) {
        wchar_t* wtxt0 = utf8_to_wide(csv_data[r][0]);
        LVITEMW lviw; lviw.mask = LVIF_TEXT; lviw.iItem = r; lviw.iSubItem = 0; lviw.pszText = wtxt0 ? wtxt0 : L"?";
        SendMessageW(hEditorListView, LVM_INSERTITEMW, 0, (LPARAM)&lviw);
        if (wtxt0) free(wtxt0);
        for(int c=1; c<csv_col_count; c++) {
            wchar_t* wtxt = utf8_to_wide(csv_data[r][c]);
            lviw.iSubItem = c; lviw.pszText = wtxt ? wtxt : L"?";
            SendMessageW(hEditorListView, LVM_SETITEMTEXTW, r, (LPARAM)&lviw);
            if (wtxt) free(wtxt);
        }
    }
}

void save_csv_memory() {
    FILE* f = fopen(mapping_file_path, "wb");
    if(!f) { MessageBox(NULL, "Cannot write CSV file! Is it open in Excel?", "Error", MB_ICONERROR); return; }
    fprintf(f, "\xEF\xBB\xBF");
    for(int c=0; c<csv_col_count; c++) fprintf(f, "\"%s\"%s", csv_headers[c], (c == csv_col_count-1) ? "" : ",");
    fprintf(f, "\n");
    for(int r=0; r<csv_row_count; r++) {
        for(int c=0; c<csv_col_count; c++) fprintf(f, "\"%s\"%s", csv_data[r][c], (c == csv_col_count-1) ? "" : ",");
        fprintf(f, "\n");
    }
    fclose(f);
    MessageBox(hEditorWnd, "Mapping dictionary saved directly into the CSV!", "Success", MB_ICONINFORMATION);
}

LRESULT CALLBACK EditorProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    // Exact same Editor Logic as previous version
    switch (uMsg) {
        case WM_CREATE:
            hEditorListView = CreateWindow(WC_LISTVIEW, "", WS_VISIBLE | WS_CHILD | WS_BORDER | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
                                     10, 10, 760, 400, hwnd, (HMENU)101, NULL, NULL);
            ListView_SetExtendedListViewStyle(hEditorListView, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
            CreateWindow("STATIC", "Modify Target DB Code for selected item:", WS_VISIBLE | WS_CHILD, 10, 420, 250, 20, hwnd, NULL, NULL, NULL);
            hEditCode = CreateWindow("EDIT", "", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 270, 418, 200, 24, hwnd, NULL, NULL, NULL);
            CreateWindow("BUTTON", "Apply Row Update", WS_VISIBLE | WS_CHILD, 480, 415, 130, 30, hwnd, (HMENU)102, NULL, NULL);
            CreateWindow("BUTTON", "SAVE TO CSV", WS_VISIBLE | WS_CHILD, 620, 415, 150, 30, hwnd, (HMENU)103, NULL, NULL);
            load_csv_memory(); populate_editor_listview();
            break;
        case WM_NOTIFY:
            if (((LPNMHDR)lParam)->idFrom == 101 && ((LPNMHDR)lParam)->code == LVN_ITEMCHANGED) {
                LPNMLISTVIEW pnmv = (LPNMLISTVIEW)lParam;
                if (pnmv->uNewState & LVIS_SELECTED) {
                    selected_row = pnmv->iItem;
                    if(selected_row >= 0 && selected_row < csv_row_count && csv_col_count > 0) {
                        wchar_t* wtxt = utf8_to_wide(csv_data[selected_row][csv_col_count - 1]);
                        SendMessageW(hEditCode, WM_SETTEXT, 0, (LPARAM)wtxt);
                        if (wtxt) free(wtxt);
                    }
                }
            } break;
        case WM_COMMAND:
            if (LOWORD(wParam) == 102) { 
                if (selected_row >= 0 && selected_row < csv_row_count && csv_col_count > 0) {
                    wchar_t wbuf[256]; SendMessageW(hEditCode, WM_GETTEXT, 256, (LPARAM)wbuf);
                    char* ubuf = wide_to_utf8(wbuf);
                    if (ubuf) {
                        free(csv_data[selected_row][csv_col_count - 1]);
                        csv_data[selected_row][csv_col_count - 1] = strdup(ubuf);
                        LVITEMW lviw; lviw.mask = LVIF_TEXT; lviw.iItem = selected_row; lviw.iSubItem = csv_col_count - 1; lviw.pszText = wbuf;
                        SendMessageW(hEditorListView, LVM_SETITEMTEXTW, selected_row, (LPARAM)&lviw);
                        free(ubuf);
                        MessageBox(hwnd, "Row updated! Don't forget to push 'Save to CSV'.", "Updated", MB_ICONINFORMATION);
                    }
                }
            }
            if (LOWORD(wParam) == 103) save_csv_memory();
            break;
        case WM_CLOSE: DestroyWindow(hwnd); hEditorWnd = NULL; break;
    }
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

// --- Main Stock Logic (Omnichannel) ---
int process_stock() {
    mapping_count = 0; deduction_count = 0;
    char log_buf[256];

    FILE* fMap = fopen(mapping_file_path, "r");
    if (!fMap) { add_log("Error: Cannot open mapping file!"); return 0; }
    
    char line[MAX_LINE];
    int map_sku_col = -1, map_db_col = -1;
    if (fgets(line, MAX_LINE, fMap)) {
        char* header = line;
        if ((unsigned char)header[0] == 0xEF && (unsigned char)header[1] == 0xBB && (unsigned char)header[2] == 0xBF) header += 3;
        char buf[MAX_LINE]; strcpy(buf, header); char* cursor = buf;
        for (int i = 0; i < 50; i++) {
            char* val = next_field(&cursor); if (!val) break; trim(val);
            if (strcmp(val, "SKU") == 0 || strcmp(val, "Key") == 0) map_sku_col = i;
            if (strcmp(val, "DB Code") == 0 || strcmp(val, "code") == 0) map_db_col = i;
        }
    }
    if (map_sku_col == -1) map_sku_col = 0;
    if (map_db_col == -1) map_db_col = 1;

    while (fgets(line, MAX_LINE, fMap) && mapping_count < MAX_ITEMS) {
        char buf[MAX_LINE]; strcpy(buf, line); char* cursor = buf;
        char* key = NULL; char* sku = NULL;
        for (int i = 0; i < 50; i++) {
            char* val = next_field(&cursor); if (!val) break;
            if (i == map_sku_col) key = val; if (i == map_db_col) sku = val;
        }
        if (key && sku) {
            trim(key); trim(sku);
            strncpy(mappings[mapping_count].key, key, ID_LEN - 1);
            strncpy(mappings[mapping_count].sku, sku, SKU_LEN - 1);
            mapping_count++;
        }
    }
    fclose(fMap);
    sprintf(log_buf, ">> Loaded %d mappings from %s.", mapping_count, mapping_file_path); add_log(log_buf);

    char actual_sales_path[MAX_PATH];
    strcpy(actual_sales_path, sales_file_path);

    if (strstr(sales_file_path, ".xlsx") != NULL || strstr(sales_file_path, ".xls") != NULL) {
        add_log(">> Excel file detected. Converting to CSV explicitly via Stock Engine...");
        char cmd[2048];
        snprintf(cmd, sizeof(cmd), "stock_engine.exe convert \"%s\"", sales_file_path);
        
        STARTUPINFO si = { sizeof(si) };
        PROCESS_INFORMATION pi;
        if (CreateProcessA(NULL, cmd, NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi)) {
            WaitForSingleObject(pi.hProcess, 15000); // 15 seconds max timeout
            CloseHandle(pi.hProcess);
            CloseHandle(pi.hThread);
        }
        strcat(actual_sales_path, ".csv");
        add_log(">> Conversion complete.");
    }

    FILE* fSales = fopen(actual_sales_path, "r");
    if (!fSales) { add_log("Error: Cannot open sales file!"); return 0; }
    
    int sku_col = -1; int qty_col = -1;
    PlatformConfig* cfg = &platforms[current_platform_idx];

    if (fgets(line, MAX_LINE, fSales)) {
        char buf[MAX_LINE]; strcpy(buf, line); char* cursor = buf;
        for (int i = 0; i < 100; i++) {
            char* val = next_field(&cursor); if (!val) break; trim(val);
            if (strcmp(val, cfg->id_col) == 0) sku_col = i;
            if (strcmp(val, cfg->qty_col) == 0) qty_col = i;
        }
    }
    if (sku_col == -1 || qty_col == -1) {
        sprintf(log_buf, "Error: Could not find columns '%s' or '%s' in sales file!", cfg->id_col, cfg->qty_col);
        add_log(log_buf); fclose(fSales); return 0;
    }

    int order_count = 0;
    int missing_count = 0;
    
    ListView_DeleteAllItems(hResultListView); // Clear previous results

    while (fgets(line, MAX_LINE, fSales)) {
        char buf[MAX_LINE]; strcpy(buf, line); char* cursor = buf;
        char* key = NULL; char* qty_str = NULL;
        for (int i = 0; i < 100; i++) {
            char* val = next_field(&cursor); if (!val) break;
            if (i == sku_col) key = val; if (i == qty_col) qty_str = val;
        }
        if (key && qty_str && strlen(key) > 0) {
            trim(key); trim(qty_str);
            int qty = atoi(qty_str); if (qty <= 0 || qty > 100000) continue; 
            
            const char* sku_raw = find_sku(key);
            if (sku_raw != NULL && strcmp(sku_raw, "MISSING_IN_DB") != 0 && strlen(sku_raw) > 0) {
                char multi_sku[SKU_LEN]; strncpy(multi_sku, sku_raw, SKU_LEN - 1); multi_sku[SKU_LEN - 1] = '\0';
                
                char* token = strtok(multi_sku, "+");
                while (token != NULL) {
                    trim(token);
                    if (strlen(token) > 0) {
                        int found = 0;
                        for (int i = 0; i < deduction_count; i++) {
                            if (strcmp(deductions[i].sku, token) == 0) { deductions[i].deduction += qty; found = 1; break; }
                        }
                        if (!found && deduction_count < MAX_ITEMS) {
                            strncpy(deductions[deduction_count].sku, token, SKU_LEN - 1);
                            deductions[deduction_count].deduction = qty; deduction_count++;
                        }
                    }
                    token = strtok(NULL, "+");
                } order_count++;
            } else {
                // ITEM NOT FOUND OR MISSING
                char err_msg[512];
                if (sku_raw && strcmp(sku_raw, "MISSING_IN_DB") == 0) {
                    snprintf(err_msg, sizeof(err_msg), "[NOT BOUND IN APP] %s (Qty: %d) -> You must bind this in Editor!", key, qty);
                } else {
                    snprintf(err_msg, sizeof(err_msg), "[NOT IN DICTIONARY] %s (Qty: %d)", key, qty);
                }
                add_log(err_msg);
                missing_count++;
            }
        }
    }
    fclose(fSales);
    sprintf(log_buf, ">> Valid orders matched: %d. Found %d unbound/missing items.", order_count, missing_count); add_log(log_buf);
    
    if (deduction_count == 0) {
         add_log("No items matched for deduction. Please review your mapping.");
         EnableWindow(hBtnSync, FALSE);
    } else {
        // Populate GUI List View for Deductions
        for(int i = 0; i < deduction_count; i++) {
            LVITEMA lvi; lvi.mask = LVIF_TEXT; lvi.iItem = i; lvi.iSubItem = 0; lvi.pszText = deductions[i].sku;
            SendMessageA(hResultListView, LVM_INSERTITEMA, 0, (LPARAM)&lvi);
            char qty_s[32]; sprintf(qty_s, "%d", deductions[i].deduction);
            lvi.iSubItem = 1; lvi.pszText = qty_s;
            SendMessageA(hResultListView, LVM_SETITEMTEXTA, i, (LPARAM)&lvi);
        }
        char success_msg[128];
        sprintf(success_msg, "Preview loaded! %d unique Sub-Items calculated.", deduction_count);
        add_log(success_msg);
        EnableWindow(hBtnSync, TRUE); // Enable Sync
    }
    
    return 1;
}

void execute_sync() {
    if (deduction_count == 0) return;
    FILE* json_file = fopen("google_payload.json", "w");
    fprintf(json_file, "[\n");
    for (int i = 0; i < deduction_count; i++) {
        fprintf(json_file, "  {\"sku\": \"%s\", \"deduction\": %d}%s\n", deductions[i].sku, deductions[i].deduction, (i == deduction_count - 1) ? "" : ",");
    }
    fprintf(json_file, "]\n"); fclose(json_file);

    PlatformConfig* cfg = &platforms[current_platform_idx];
    char command[MAX_LINE];
    snprintf(command, sizeof(command), 
        "-NoProfile -WindowStyle Hidden -Command \"Invoke-RestMethod -Uri '%s' -Method Post -ContentType 'application/json' -InFile '.\\google_payload.json'\"", 
        strlen(cfg->webhook) > 5 ? cfg->webhook : google_webhook_url);
    ShellExecuteA(NULL, "open", "powershell.exe", command, NULL, SW_HIDE);
    add_log(">> SUCCESS: Sync fired via Webhook in Background!");
    EnableWindow(hBtnSync, FALSE); // Disable until next process
}

// --- Main Window ---
void SelectFile(HWND hwnd, char* path, HWND hLabel) {
    OPENFILENAME ofn; ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn); ofn.hwndOwner = hwnd; ofn.lpstrFile = path; ofn.nMaxFile = MAX_PATH;
    ofn.lpstrFilter = "CSV/Excel\0*.*\0"; ofn.nFilterIndex = 1; ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_NOCHANGEDIR;
    if (GetOpenFileName(&ofn)) SetWindowText(hLabel, path);
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
        case WM_CREATE:
            load_config();
            CreateWindow("STATIC", "Select Platform Context:", WS_VISIBLE | WS_CHILD, 20, 20, 160, 20, hwnd, NULL, NULL, NULL);
            hComboPlatform = CreateWindow(WC_COMBOBOX, "", CBS_DROPDOWNLIST | WS_CHILD | WS_VISIBLE, 180, 15, 230, 200, hwnd, (HMENU)5, NULL, NULL);
            for (int i=0; i<platform_count; i++) SendMessage(hComboPlatform, CB_ADDSTRING, 0, (LPARAM)platforms[i].name);
            SendMessage(hComboPlatform, CB_SETCURSEL, 0, 0);

            CreateWindow("STATIC", "1. Dictionary Mapping File:", WS_VISIBLE | WS_CHILD, 20, 50, 200, 20, hwnd, NULL, NULL, NULL);
            CreateWindow("BUTTON", "Edit Mapping", WS_VISIBLE | WS_CHILD, 220, 45, 110, 25, hwnd, (HMENU)4, NULL, NULL);
            CreateWindow("BUTTON", "Gen. Map (Auto)", WS_VISIBLE | WS_CHILD, 340, 45, 130, 25, hwnd, (HMENU)8, NULL, NULL);
            hMapPath = CreateWindow("STATIC", platforms[0].map_file, WS_VISIBLE | WS_CHILD | SS_LEFTNOWORDWRAP, 20, 75, 300, 20, hwnd, NULL, NULL, NULL);
            strcpy(mapping_file_path, platforms[0].map_file);
            CreateWindow("BUTTON", "Browse...", WS_VISIBLE | WS_CHILD, 330, 70, 80, 25, hwnd, (HMENU)1, NULL, NULL);

            CreateWindow("STATIC", "2. Sales Order File (CSV):", WS_VISIBLE | WS_CHILD, 20, 105, 250, 20, hwnd, NULL, NULL, NULL);
            hSalesPath = CreateWindow("STATIC", "Select a file...", WS_VISIBLE | WS_CHILD | SS_LEFTNOWORDWRAP, 20, 130, 300, 20, hwnd, NULL, NULL, NULL);
            CreateWindow("BUTTON", "Browse...", WS_VISIBLE | WS_CHILD, 330, 125, 80, 25, hwnd, (HMENU)2, NULL, NULL);

            hBtnProcess = CreateWindow("BUTTON", "1. ANALYZE DEDUCTIONS", WS_VISIBLE | WS_CHILD, 20, 165, 190, 40, hwnd, (HMENU)3, NULL, NULL);
            hBtnSync = CreateWindow("BUTTON", "2. SYNC TO GOOGLE SHEETS", WS_VISIBLE | WS_CHILD | WS_DISABLED, 220, 165, 190, 40, hwnd, (HMENU)6, NULL, NULL);
            
            CreateWindow("STATIC", "Preview Output Payload:", WS_VISIBLE | WS_CHILD, 20, 215, 200, 20, hwnd, NULL, NULL, NULL);
            hResultListView = CreateWindow(WC_LISTVIEW, "", WS_VISIBLE | WS_CHILD | WS_BORDER | LVS_REPORT | LVS_SINGLESEL,
                                     20, 235, 390, 150, hwnd, (HMENU)7, NULL, NULL);
            ListView_SetExtendedListViewStyle(hResultListView, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
            LVCOLUMNA lvc; lvc.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
            lvc.iSubItem = 0; lvc.pszText = "Central DB Code (SKU)"; lvc.cx = 250; SendMessageA(hResultListView, LVM_INSERTCOLUMNA, 0, (LPARAM)&lvc);
            lvc.iSubItem = 1; lvc.pszText = "Qty to Deduct"; lvc.cx = 120; SendMessageA(hResultListView, LVM_INSERTCOLUMNA, 1, (LPARAM)&lvc);

            CreateWindow("STATIC", "Console Event Log (Errors/Missing ID warnings):", WS_VISIBLE | WS_CHILD, 20, 390, 480, 20, hwnd, NULL, NULL, NULL);
            hStatus = CreateWindow("EDIT", "", WS_VISIBLE | WS_CHILD | WS_VSCROLL | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY, 20, 410, 480, 120, hwnd, NULL, NULL, NULL);
            add_log("Omnichannel Stock Engine v3 Loaded.");
            break;
        case WM_COMMAND:
            if (HIWORD(wParam) == CBN_SELCHANGE && LOWORD(wParam) == 5) {
                current_platform_idx = SendMessage(hComboPlatform, CB_GETCURSEL, 0, 0);
                strcpy(mapping_file_path, platforms[current_platform_idx].map_file);
                SetWindowText(hMapPath, mapping_file_path);
                char inf[128]; sprintf(inf, "Switched context to %s.", platforms[current_platform_idx].name); add_log(inf);
            }
            if (LOWORD(wParam) == 1) SelectFile(hwnd, mapping_file_path, hMapPath);
            if (LOWORD(wParam) == 2) SelectFile(hwnd, sales_file_path, hSalesPath);
            if (LOWORD(wParam) == 4) { 
                if (!hEditorWnd) hEditorWnd = CreateWindow("MappingEditorClass", "Omnichannel Dictionary Editor", WS_VISIBLE | WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 800, 500, hwnd, NULL, GetModuleHandle(NULL), NULL);
                else BringWindowToTop(hEditorWnd);
            }
            if (LOWORD(wParam) == 3) {
                if (strlen(mapping_file_path) == 0 || strlen(sales_file_path) == 0) MessageBox(hwnd, "Please select both files first!", "Warning", MB_ICONWARNING);
                else {
                    add_log("\n--- Analyzing Orders ---");
                    process_stock();
                }
            }
            if (LOWORD(wParam) == 6) {
                execute_sync();
            }
            if (LOWORD(wParam) == 8) {
                add_log(">> Firing Auto-Map Generator Engine...");
                STARTUPINFO si = { sizeof(si) }; PROCESS_INFORMATION pi;
                if (CreateProcessA(NULL, "stock_engine.exe map", NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi)) {
                    WaitForSingleObject(pi.hProcess, 30000); 
                    CloseHandle(pi.hProcess); CloseHandle(pi.hThread);
                    MessageBox(hwnd, "Dictionary updated internally! Press ANALYZE again to refresh matches.", "Auto-Map Complete", MB_ICONINFORMATION);
                    add_log(">> Dictionary matrix regenerated successfully!");
                } else {
                    add_log("Error: 'stock_engine.exe' not found! Missing Python payload module.");
                }
            }
            break;
        case WM_DESTROY: PostQuitMessage(0); return 0;
    }
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    INITCOMMONCONTROLSEX icex; icex.dwSize = sizeof(INITCOMMONCONTROLSEX); icex.dwICC = ICC_LISTVIEW_CLASSES;
    InitCommonControlsEx(&icex);
    
    WNDCLASS wc = {0}, wcEd = {0};
    wc.lpfnWndProc = WindowProc; wc.hInstance = hInstance; wc.lpszClassName = "OmniStockSyncManager";
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW); wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClass(&wc);

    wcEd.lpfnWndProc = EditorProc; wcEd.hInstance = hInstance; wcEd.lpszClassName = "MappingEditorClass";
    wcEd.hbrBackground = (HBRUSH)(COLOR_BTNFACE); wcEd.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClass(&wcEd);

    hMainWnd = CreateWindowEx(0, "OmniStockSyncManager", "Smart Stock Sync - Omnichannel Edition", WS_OVERLAPPEDWINDOW & ~WS_THICKFRAME & ~WS_MAXIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, 550, 580, NULL, NULL, hInstance, NULL);
    if (!hMainWnd) return 0;
    ShowWindow(hMainWnd, nCmdShow);
    
    MSG msg = {0};
    while (GetMessage(&msg, NULL, 0, 0)) { TranslateMessage(&msg); DispatchMessage(&msg); }
    return 0;
}
