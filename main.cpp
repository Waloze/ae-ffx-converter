#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN

#include <windows.h>
#include <commdlg.h>
#include <shlobj.h>
#include <string>
#include <vector>
#include <map>
#include <fstream>
#include <sstream>
#include <set>
#include <algorithm>

#include <shellapi.h>
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")

// ── Version Table ─────────────────────────────────────────────────────────────

struct Version { BYTE major, minor; };

const std::vector<std::wstring> VERSION_ORDER = {
    L"CS6", L"CC (12.x)", L"CC 2014", L"CC 2015.3", L"CC 2017",
    L"CC 2018", L"CC 2019", L"2020", L"2021", L"2022", L"2023", L"2024", L"2025"
};

const std::map<std::wstring, Version> VERSIONS = {
    {L"CS6",       {0x51, 0x08}},
    {L"CC (12.x)", {0x56, 0x04}},
    {L"CC 2014",   {0x57, 0x01}},
    {L"CC 2015.3", {0x57, 0x03}},
    {L"CC 2017",   {0x58, 0x0A}},
    {L"CC 2018",   {0x5C, 0x0E}},
    {L"CC 2019",   {0x5D, 0x05}},
    {L"2020",      {0x5D, 0x06}},
    {L"2021",      {0x5D, 0x1D}},
    {L"2022",      {0x5D, 0x2B}},
    {L"2023",      {0x5E, 0x09}},
    {L"2024",      {0x5F, 0x06}},
    {L"2025",      {0x60, 0x06}},
};

const int OFFSET_MAJOR = 0x1B;
const int OFFSET_MINOR = 0x1F;

// ── Colors ────────────────────────────────────────────────────────────────────

#define COL_BG       RGB(0x1e, 0x1e, 0x2e)
#define COL_SURFACE  RGB(0x2a, 0x2a, 0x3e)
#define COL_ACCENT   RGB(0x7c, 0x6a, 0xf7)
#define COL_TEXT     RGB(0xcd, 0xd6, 0xf4)
#define COL_SUBTEXT  RGB(0x6c, 0x70, 0x86)
#define COL_SUCCESS  RGB(0xa6, 0xe3, 0xa1)
#define COL_ERROR    RGB(0xf3, 0x8b, 0xa8)
#define COL_BORDER   RGB(0x45, 0x47, 0x5a)

// ── Control IDs ───────────────────────────────────────────────────────────────

#define ID_BTN_SELECT    101
#define ID_BTN_CLEAR     102
#define ID_BTN_OUTPUT    103
#define ID_BTN_CONVERT   104
#define ID_COMBO_VERSION 105
#define ID_LBL_FILES     106
#define ID_LBL_DETECTED  107
#define ID_LBL_OUTPUT    108
#define ID_LBL_STATUS    109

// ── Globals ───────────────────────────────────────────────────────────────────

HWND hWnd;
HWND hLblFiles, hLblDetected, hLblOutput, hLblStatus;
HWND hCombo;
HWND hBtnSelect, hBtnClear, hBtnOutput, hBtnConvert;

HBRUSH hBrushBg, hBrushSurface, hBrushAccent, hBrushBorder;
HFONT hFontTitle, hFontNormal, hFontSmall, hFontBold;

std::vector<std::wstring> selectedFiles;
std::wstring outputFolder = L"";

// ── Core Logic ────────────────────────────────────────────────────────────────

std::wstring detectVersion(const std::vector<BYTE>& data) {
    if (data.size() <= OFFSET_MINOR) return L"Unknown";
    BYTE major = data[OFFSET_MAJOR];
    BYTE minor = data[OFFSET_MINOR];
    for (auto& name : VERSION_ORDER) {
        auto it = VERSIONS.find(name);
        if (it != VERSIONS.end() && it->second.major == major && it->second.minor == minor)
            return name;
    }
    wchar_t buf[64];
    swprintf_s(buf, L"Unknown (0x%02X, 0x%02X)", major, minor);
    return buf;
}

bool convertFile(const std::wstring& inPath, const std::wstring& outPath, const std::wstring& target) {
    std::ifstream in(inPath.c_str(), std::ios::binary);
    if (!in) return false;
    std::vector<BYTE> data((std::istreambuf_iterator<char>(in)), {});
    in.close();

    if (data.size() <= OFFSET_MINOR) return false;
    auto it = VERSIONS.find(target);
    if (it == VERSIONS.end()) return false;

    data[OFFSET_MAJOR] = it->second.major;
    data[OFFSET_MINOR] = it->second.minor;

    std::ofstream out(outPath.c_str(), std::ios::binary);
    if (!out) return false;
    out.write(reinterpret_cast<char*>(data.data()), data.size());
    return true;
}

std::wstring safeVersionName(const std::wstring& v) {
    std::wstring s;
    for (wchar_t c : v) {
        if (c == L' ') s += L'_';
        else if (c == L'(' || c == L')' || c == L'.') continue;
        else s += c;
    }
    return s;
}

std::wstring getFilename(const std::wstring& path) {
    size_t pos = path.find_last_of(L"\\/");
    return pos == std::wstring::npos ? path : path.substr(pos + 1);
}

std::wstring getDirectory(const std::wstring& path) {
    size_t pos = path.find_last_of(L"\\/");
    return pos == std::wstring::npos ? L"." : path.substr(0, pos);
}

std::wstring stripExt(const std::wstring& filename) {
    size_t pos = filename.find_last_of(L'.');
    return pos == std::wstring::npos ? filename : filename.substr(0, pos);
}

// ── UI Helpers ────────────────────────────────────────────────────────────────

void updateFileLabels() {
    if (selectedFiles.empty()) {
        SetWindowText(hLblFiles, L"No files selected");
        SetWindowText(hLblDetected, L"—");
        return;
    }

    wchar_t buf[256];
    swprintf_s(buf, L"%d file(s) selected", (int)selectedFiles.size());
    SetWindowText(hLblFiles, buf);

    // Detect versions
    std::set<std::wstring> versions;
    for (auto& path : selectedFiles) {
        std::ifstream f(path.c_str(), std::ios::binary);
        if (f) {
            std::vector<BYTE> data((std::istreambuf_iterator<char>(f)), {});
            versions.insert(detectVersion(data));
        }
    }
    std::wstring detected;
    for (auto& v : versions) {
        if (!detected.empty()) detected += L", ";
        detected += v;
    }
    SetWindowText(hLblDetected, detected.c_str());
}

void selectFiles() {
    // Build a wide multi-select filter buffer
    wchar_t fileBuffer[32768] = {0};

    OPENFILENAMEW ofn = {};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hWnd;
    ofn.lpstrFilter = L"After Effects Presets (*.ffx)\0*.ffx\0All Files (*.*)\0*.*\0";
    ofn.lpstrFile = fileBuffer;
    ofn.nMaxFile = sizeof(fileBuffer) / sizeof(wchar_t);
    ofn.Flags = OFN_ALLOWMULTISELECT | OFN_EXPLORER | OFN_FILEMUSTEXIST;
    ofn.lpstrTitle = L"Select FFX Files";

    if (GetOpenFileNameW(&ofn)) {
        // Parse multi-select result
        // If multiple files: first string is directory, then filenames separated by nulls
        wchar_t* p = fileBuffer;
        std::wstring dir = p;
        p += dir.size() + 1;

        if (*p == L'\0') {
            // Single file selected
            selectedFiles.push_back(dir);
        } else {
            // Multiple files
            while (*p != L'\0') {
                std::wstring filename = p;
                selectedFiles.push_back(dir + L"\\" + filename);
                p += filename.size() + 1;
            }
        }
        updateFileLabels();
    }
}

void browseOutput() {
    wchar_t path[MAX_PATH] = {0};
    BROWSEINFOW bi = {};
    bi.hwndOwner = hWnd;
    bi.pszDisplayName = path;
    bi.lpszTitle = L"Select Output Folder";
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;

    LPITEMIDLIST pidl = SHBrowseForFolderW(&bi);
    if (pidl) {
        SHGetPathFromIDListW(pidl, path);
        CoTaskMemFree(pidl);
        outputFolder = path;
        SetWindowText(hLblOutput, path);
    }
}

void doConvert() {
    if (selectedFiles.empty()) {
        MessageBoxW(hWnd, L"Please select at least one .ffx file.", L"No files", MB_ICONWARNING);
        return;
    }

    wchar_t targetBuf[64] = {0};
    GetWindowTextW(hCombo, targetBuf, 64);
    std::wstring target = targetBuf;

    std::vector<std::wstring> results, errors;

    for (auto& f : selectedFiles) {
        std::wstring dir = outputFolder.empty() ? getDirectory(f) : outputFolder;
        std::wstring base = stripExt(getFilename(f));
        std::wstring safe = safeVersionName(target);
        std::wstring outPath = dir + L"\\" + base + L"_" + safe + L".ffx";

        // Get source version name
        std::wstring source;
        {
            std::ifstream fin(f.c_str(), std::ios::binary);
            if (fin) {
                std::vector<BYTE> data((std::istreambuf_iterator<char>(fin)), {});
                source = detectVersion(data);
            }
        }

        if (convertFile(f, outPath, target)) {
            results.push_back(getFilename(f) + L"  (" + source + L" → " + target + L")");
        } else {
            errors.push_back(getFilename(f) + L": failed to convert");
        }
    }

    // Status label
    wchar_t statusBuf[128];
    swprintf_s(statusBuf, L"✓ Done — %d file(s) converted", (int)results.size());
    SetWindowText(hLblStatus, statusBuf);

    // Build result message
    std::wstring msg = L"✓ Converted " + std::to_wstring(results.size()) + L" file(s) to " + target + L"\n\n";
    for (auto& r : results) msg += r + L"\n";
    if (!errors.empty()) {
        msg += L"\nErrors:\n";
        for (auto& e : errors) msg += e + L"\n";
    }
    MessageBoxW(hWnd, msg.c_str(), L"Done", MB_ICONINFORMATION);
}

// ── Window Procedure ──────────────────────────────────────────────────────────

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {

    case WM_DROPFILES: {
        HDROP hDrop = (HDROP)wParam;
        UINT count = DragQueryFileW(hDrop, 0xFFFFFFFF, NULL, 0);
        for (UINT i = 0; i < count; i++) {
            wchar_t path[MAX_PATH];
            DragQueryFileW(hDrop, i, path, MAX_PATH);
            std::wstring p = path;
            // Only accept .ffx files
            if (p.size() >= 4 && p.substr(p.size() - 4) == L".ffx") {
                // Avoid duplicates
                bool found = false;
                for (auto& f : selectedFiles) if (f == p) { found = true; break; }
                if (!found) selectedFiles.push_back(p);
            }
        }
        DragFinish(hDrop);
        updateFileLabels();
        return 0;
    }

    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)wParam;
        HWND hCtrl = (HWND)lParam;
        SetBkMode(hdc, TRANSPARENT);
        if (hCtrl == hLblStatus) {
            SetTextColor(hdc, COL_SUCCESS);
        } else if (hCtrl == hLblDetected) {
            SetTextColor(hdc, RGB(0xa7, 0x8b, 0xfa));
        } else if (hCtrl == hLblFiles || hCtrl == hLblOutput) {
            SetTextColor(hdc, COL_SUBTEXT);
        } else {
            SetTextColor(hdc, COL_TEXT);
        }
        return (LRESULT)hBrushBg;
    }

    case WM_CTLCOLORBTN:
        return (LRESULT)hBrushBg;

    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORLISTBOX: {
        HDC hdc = (HDC)wParam;
        SetTextColor(hdc, COL_TEXT);
        SetBkColor(hdc, COL_SURFACE);
        return (LRESULT)hBrushSurface;
    }

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case ID_BTN_SELECT:  selectFiles(); break;
        case ID_BTN_CLEAR:
            selectedFiles.clear();
            updateFileLabels();
            break;
        case ID_BTN_OUTPUT:  browseOutput(); break;
        case ID_BTN_CONVERT: doConvert(); break;
        }
        break;

    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);

        // Background
        RECT rc;
        GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, hBrushBg);

        // Title text
        SelectObject(hdc, hFontTitle);
        SetTextColor(hdc, RGB(0xa7, 0x8b, 0xfa));
        SetBkMode(hdc, TRANSPARENT);
        RECT titleRect = {20, 16, 460, 44};
        DrawTextW(hdc, L"FFX Version Converter", -1, &titleRect, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

        // Subtitle
        SelectObject(hdc, hFontSmall);
        SetTextColor(hdc, COL_SUBTEXT);
        RECT subRect = {20, 44, 460, 62};
        DrawTextW(hdc, L"Convert After Effects presets between any version. Fully offline.", -1, &subRect, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

        // File area box
        RECT boxRect = {20, 68, 460, 148};
        FrameRect(hdc, &boxRect, hBrushBorder);
        HBRUSH oldBrush = (HBRUSH)SelectObject(hdc, hBrushSurface);
        boxRect.left++; boxRect.top++; boxRect.right--; boxRect.bottom--;
        FillRect(hdc, &boxRect, hBrushSurface);
        SelectObject(hdc, oldBrush);

        // Separator lines
        HPEN hPen = CreatePen(PS_SOLID, 1, COL_BORDER);
        HPEN hOldPen = (HPEN)SelectObject(hdc, hPen);
        MoveToEx(hdc, 20, 168, NULL); LineTo(hdc, 460, 168);
        MoveToEx(hdc, 20, 230, NULL); LineTo(hdc, 460, 230);
        SelectObject(hdc, hOldPen);
        DeleteObject(hPen);

        // Labels
        SelectObject(hdc, hFontNormal);
        SetTextColor(hdc, COL_SUBTEXT);
        RECT detLbl = {20, 150, 160, 170};
        DrawTextW(hdc, L"Detected version:", -1, &detLbl, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

        SetTextColor(hdc, COL_TEXT);
        RECT targetLbl = {20, 180, 110, 200};
        DrawTextW(hdc, L"Convert to:", -1, &targetLbl, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

        RECT outputLbl = {20, 240, 120, 260};
        DrawTextW(hdc, L"Output folder:", -1, &outputLbl, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

        EndPaint(hwnd, &ps);
        break;
    }

    case WM_ERASEBKGND: {
        HDC hdc = (HDC)wParam;
        RECT rc;
        GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, hBrushBg);
        return 1;
    }

    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

// ── Custom Button ─────────────────────────────────────────────────────────────

LRESULT CALLBACK AccentBtnProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR) {
    if (msg == WM_PAINT) {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        RECT rc;
        GetClientRect(hwnd, &rc);

        bool hover = false;
        POINT pt; GetCursorPos(&pt); ScreenToClient(hwnd, &pt);
        if (PtInRect(&rc, pt)) hover = true;

        COLORREF bg = hover ? RGB(0xa7, 0x8b, 0xfa) : COL_ACCENT;
        HBRUSH br = CreateSolidBrush(bg);
        FillRect(hdc, &rc, br);
        DeleteObject(br);

        wchar_t text[64] = {0};
        GetWindowTextW(hwnd, text, 64);
        SelectObject(hdc, hFontBold);
        SetTextColor(hdc, RGB(0xff, 0xff, 0xff));
        SetBkMode(hdc, TRANSPARENT);
        DrawTextW(hdc, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        EndPaint(hwnd, &ps);
        return 0;
    }
    if (msg == WM_MOUSEMOVE || msg == WM_MOUSELEAVE) InvalidateRect(hwnd, NULL, FALSE);
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

LRESULT CALLBACK SmallBtnProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR, DWORD_PTR) {
    if (msg == WM_PAINT) {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        RECT rc;
        GetClientRect(hwnd, &rc);

        bool hover = false;
        POINT pt; GetCursorPos(&pt); ScreenToClient(hwnd, &pt);
        if (PtInRect(&rc, pt)) hover = true;

        COLORREF bg = hover ? COL_ACCENT : COL_SURFACE;
        HBRUSH br = CreateSolidBrush(bg);
        FillRect(hdc, &rc, br);
        DeleteObject(br);

        // Border
        HPEN pen = CreatePen(PS_SOLID, 1, COL_BORDER);
        HPEN oldPen = (HPEN)SelectObject(hdc, pen);
        HBRUSH oldBr = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
        Rectangle(hdc, rc.left, rc.top, rc.right, rc.bottom);
        SelectObject(hdc, oldPen);
        SelectObject(hdc, oldBr);
        DeleteObject(pen);

        wchar_t text[64] = {0};
        GetWindowTextW(hwnd, text, 64);
        SelectObject(hdc, hFontSmall);
        SetTextColor(hdc, COL_TEXT);
        SetBkMode(hdc, TRANSPARENT);
        DrawTextW(hdc, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        EndPaint(hwnd, &ps);
        return 0;
    }
    if (msg == WM_MOUSEMOVE || msg == WM_MOUSELEAVE) InvalidateRect(hwnd, NULL, FALSE);
    return DefSubclassProc(hwnd, msg, wParam, lParam);
}

#pragma comment(lib, "comctl32.lib")
#include <commctrl.h>

// ── Entry Point ───────────────────────────────────────────────────────────────

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int) {
    CoInitialize(NULL);

    // Brushes
    hBrushBg      = CreateSolidBrush(COL_BG);
    hBrushSurface = CreateSolidBrush(COL_SURFACE);
    hBrushAccent  = CreateSolidBrush(COL_ACCENT);
    hBrushBorder  = CreateSolidBrush(COL_BORDER);

    // Fonts
    hFontTitle  = CreateFontW(24, 0, 0, 0, FW_BOLD,   0,0,0, DEFAULT_CHARSET, 0,0, CLEARTYPE_QUALITY, 0, L"Segoe UI");
    hFontBold   = CreateFontW(16, 0, 0, 0, FW_BOLD,   0,0,0, DEFAULT_CHARSET, 0,0, CLEARTYPE_QUALITY, 0, L"Segoe UI");
    hFontNormal = CreateFontW(15, 0, 0, 0, FW_NORMAL, 0,0,0, DEFAULT_CHARSET, 0,0, CLEARTYPE_QUALITY, 0, L"Segoe UI");
    hFontSmall  = CreateFontW(13, 0, 0, 0, FW_NORMAL, 0,0,0, DEFAULT_CHARSET, 0,0, CLEARTYPE_QUALITY, 0, L"Segoe UI");

    // Register window class
    WNDCLASSW wc = {};
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInstance;
    wc.lpszClassName = L"FFXConverter";
    wc.hCursor       = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = hBrushBg;
    RegisterClassW(&wc);

    // Create window
    hWnd = CreateWindowExW(0, L"FFXConverter", L"FFX Version Converter",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, 500, 380,
        NULL, NULL, hInstance, NULL);

    // ── Controls ──────────────────────────────────────────────────────────────

    // File label (inside box)
    hLblFiles = CreateWindowW(L"STATIC", L"No files selected. Drag and drop or use buttons below",
        WS_CHILD | WS_VISIBLE | SS_CENTER,
        21, 88, 438, 24, hWnd, (HMENU)ID_LBL_FILES, hInstance, NULL);
    SendMessage(hLblFiles, WM_SETFONT, (WPARAM)hFontNormal, TRUE);

    // Buttons row (inside box)
    hBtnSelect = CreateWindowW(L"BUTTON", L"Select Files",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_OWNERDRAW,
        30, 116, 130, 26, hWnd, (HMENU)ID_BTN_SELECT, hInstance, NULL);
    SendMessage(hBtnSelect, WM_SETFONT, (WPARAM)hFontSmall, TRUE);
    SetWindowSubclass(hBtnSelect, SmallBtnProc, 0, 0);

    hBtnClear = CreateWindowW(L"BUTTON", L"Clear",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_OWNERDRAW,
        170, 116, 100, 26, hWnd, (HMENU)ID_BTN_CLEAR, hInstance, NULL);
    SendMessage(hBtnClear, WM_SETFONT, (WPARAM)hFontSmall, TRUE);
    SetWindowSubclass(hBtnClear, SmallBtnProc, 0, 0);

    // Detected version label
    hLblDetected = CreateWindowW(L"STATIC", L"—",
        WS_CHILD | WS_VISIBLE,
        160, 150, 280, 20, hWnd, (HMENU)ID_LBL_DETECTED, hInstance, NULL);
    SendMessage(hLblDetected, WM_SETFONT, (WPARAM)hFontSmall, TRUE);

    // Target version combobox
    hCombo = CreateWindowW(L"COMBOBOX", NULL,
        WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_VSCROLL,
        130, 176, 160, 300, hWnd, (HMENU)ID_COMBO_VERSION, hInstance, NULL);
    SendMessage(hCombo, WM_SETFONT, (WPARAM)hFontNormal, TRUE);
    for (auto& v : VERSION_ORDER)
        SendMessage(hCombo, CB_ADDSTRING, 0, (LPARAM)v.c_str());
    // Default to 2020
    SendMessage(hCombo, CB_SELECTSTRING, -1, (LPARAM)L"2020");

    // Output folder label + browse button
    hLblOutput = CreateWindowW(L"STATIC", L"Same as input",
        WS_CHILD | WS_VISIBLE | SS_ENDELLIPSIS,
        140, 240, 240, 20, hWnd, (HMENU)ID_LBL_OUTPUT, hInstance, NULL);
    SendMessage(hLblOutput, WM_SETFONT, (WPARAM)hFontSmall, TRUE);

    hBtnOutput = CreateWindowW(L"BUTTON", L"Browse",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_OWNERDRAW,
        390, 236, 70, 26, hWnd, (HMENU)ID_BTN_OUTPUT, hInstance, NULL);
    SendMessage(hBtnOutput, WM_SETFONT, (WPARAM)hFontSmall, TRUE);
    SetWindowSubclass(hBtnOutput, SmallBtnProc, 0, 0);

    // Convert button
    hBtnConvert = CreateWindowW(L"BUTTON", L"Convert",
        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | BS_OWNERDRAW,
        20, 278, 440, 40, hWnd, (HMENU)ID_BTN_CONVERT, hInstance, NULL);
    SendMessage(hBtnConvert, WM_SETFONT, (WPARAM)hFontBold, TRUE);
    SetWindowSubclass(hBtnConvert, AccentBtnProc, 0, 0);

    // Status label
    hLblStatus = CreateWindowW(L"STATIC", L"",
        WS_CHILD | WS_VISIBLE | SS_CENTER,
        20, 326, 440, 20, hWnd, (HMENU)ID_LBL_STATUS, hInstance, NULL);
    SendMessage(hLblStatus, WM_SETFONT, (WPARAM)hFontSmall, TRUE);

    // ── Show ──────────────────────────────────────────────────────────────────

    ShowWindow(hWnd, SW_SHOW);
    UpdateWindow(hWnd);
    DragAcceptFiles(hWnd, TRUE);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    CoUninitialize();
    return 0;
}