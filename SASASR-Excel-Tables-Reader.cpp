#include <iostream>
#include <xlnt/xlnt.hpp>
#include <windows.h>
#include <vector>
#include <string>
    #pragma comment(lib, "xlnt.lib")
using std::string, std::vector;
    struct TableEntry {
        bool hasDate;
        string date;
        int position;
        string country;
        string platform;
        string name;
        string time;
        string racer;
        string map;
    };
class Table {
public:
    vector<TableEntry> entries;
    string name;
        Table() {}
    void clear() {
        entries.clear();
    }
};
    int AAToNumber(const std::string& label) {
        int number = 0;
        int length = label.length();
            for (int i = 0; i < length; ++i) {
                char ch = label[i];
                if (std::isalpha(ch)) {
                    number = number * 26 + (std::toupper(ch) - 'A' + 1);
                } else {
                    throw std::invalid_argument("Invalid character in column label.");
                }
            }
        return number;
    }
    xlnt::workbook wb;
void processSpreadsheet(Table& table, xlnt::worksheet& ws, string x1s, int x2, int y1, int y2, int order[7]) {
    int x1 = AAToNumber(x1s) - 1;
    x2 += x1 - 1;
    y1 -= 1;
    y2 = y1 + y2 - 1;
    std::cout << "Processing spreadsheet" << std::endl;
    for (int row_index = y1; row_index <= y2; ++row_index) {
        TableEntry entry;
        int i = 0;
            for (int col_index = x1; col_index <= x2; ++col_index) {
                auto cell = ws.cell(xlnt::cell_reference(col_index + 1, row_index + 1));
                    if (i == order[0]) {
                        if (cell.has_format() && (cell.data_type() == xlnt::cell::type::date || cell.data_type() == xlnt::cell::type::number)) {
                            auto nf = cell.number_format();
                            entry.hasDate = true;
                            entry.date = nf.format(cell.value<double>(), wb.base_date());
                        }
                    } else if (i == order[1]) {
                        entry.position = cell.value<int>();
                    } else if (i == order[2]) {
                        entry.country = cell.to_string();
                    } else if (i == order[3]) {
                        entry.platform = cell.to_string();
                    } else if (i == order[4]) {
                        entry.name = cell.to_string();
                    } else if (i == order[5]) {
                        entry.time = cell.to_string();
                    } else if (i == order[6]) {
                        entry.racer = cell.to_string();
                    } else if (i == order[7]) {
                        entry.map = cell.to_string();
                    }
                ++i;
            }
        table.entries.push_back(entry);
    }
    std::cout << "Processing complete" << std::endl;
}
    #pragma warning(push)
    #pragma warning(disable : C2362)
    #pragma warning(pop)
int main() {
    try {
        wb.load("leaderboards.xlsx");
    } catch (...) {
        std::cout << "couldn't find leaderboards.xlsx, you can download it at:\nhttps://docs.google.com/spreadsheets/d/1qrvwTB2DsdGWeVcdSRq5duEa7Hep8MJ9_IxeCj_9sOI/edit?gid=0#gid=0";
        while (1);
    }
    auto ws = wb.active_sheet();
        std::vector<int> recent({ 0, 1, 2, 3, 4, 5, 7, 6 });
            int* lRecent = recent.data();
        std::vector<int> rank8({ 7, 0, 1, 2, 3, 4, 5, 6 });
            int* lRank8 = rank8.data();
    Table table;
    processSpreadsheet(table, ws, "B", 7, 4, 147, lRecent); // Recent Runs
        for (const auto& entry : table.entries) {
            std::cout << "Entry: " << entry.position << " " << entry.name << " - " << entry.time << std::endl;
            std::cout << entry.map << std::endl;
        }
        table.clear();
    ws = wb.sheet_by_index(1);  //Trial No Major Glitch
    processSpreadsheet(table, ws, "B", 6,  8, 8, lRank8); // Whale Lagoon
    processSpreadsheet(table, ws, "I", 6,  8, 8, lRank8); // Icicle Valley
    processSpreadsheet(table, ws, "P", 6,  8, 8, lRank8); // Roulette Road
    processSpreadsheet(table, ws, "W", 6,  8, 8, lRank8); // Sunshine Tour
        processSpreadsheet(table, ws, "B", 6, 19, 8, lRank8); // Shibuya Downtown
        processSpreadsheet(table, ws, "I", 6, 19, 8, lRank8); // Outer Forest
        processSpreadsheet(table, ws, "P", 6, 19, 8, lRank8); // Turbine Loop
        processSpreadsheet(table, ws, "W", 6, 19, 8, lRank8); // Treetops
    processSpreadsheet(table, ws, "B", 6, 30, 8, lRank8); // Rampart Road
    processSpreadsheet(table, ws, "I", 6, 30, 8, lRank8); // Dark Arsenal
    processSpreadsheet(table, ws, "P", 6, 30, 8, lRank8); // Jump Parade
    processSpreadsheet(table, ws, "W", 6, 30, 8, lRank8); // Pinball Highway
        processSpreadsheet(table, ws, "B", 6, 41, 8, lRank8); // Sewer Scrapes
        processSpreadsheet(table, ws, "I", 6, 41, 8, lRank8); // Lost Palace
        processSpreadsheet(table, ws, "P", 6, 41, 8, lRank8); // Sandy Drifts
        processSpreadsheet(table, ws, "W", 6, 41, 8, lRank8); // Rokkaku Hill
    processSpreadsheet(table, ws, "B", 6, 52, 8, lRank8); // Rocky-Coaster
    processSpreadsheet(table, ws, "I", 6, 52, 8, lRank8); // Highway Zero
    processSpreadsheet(table, ws, "P", 6, 52, 8, lRank8); // Deadly Route
    processSpreadsheet(table, ws, "W", 6, 52, 8, lRank8); // Ocean Ruin
        processSpreadsheet(table, ws, "B", 6, 63, 8, lRank8); // Bingo Party
        processSpreadsheet(table, ws, "I", 6, 63, 8, lRank8); // Lava Lair
        processSpreadsheet(table, ws, "P", 6, 63, 8, lRank8); // Monkey Target
        processSpreadsheet(table, ws, "W", 6, 63, 8, lRank8); // Thunder Deck
    processSpreadsheet(table, ws, "B", 6, 74, 8, lRank8); // Egg Hangar
    processSpreadsheet(table, ws, "I", 4, 74, 8, lRank8); // WR Tally
        for (const auto& entry : table.entries) {
            std::cout << "Entry: " << entry.position << " " << entry.name << " - " << entry.time << std::endl;
            std::cout << entry.racer << std::endl;
        }
        table.clear();
    ws = wb.sheet_by_index(2); // Major Glitches
    processSpreadsheet(table, ws, "B", 6, 8, 8, lRank8); // Whale Lagoon
    processSpreadsheet(table, ws, "I", 6, 8, 8, lRank8); // Shibuya Downtown
    processSpreadsheet(table, ws, "P", 6, 8, 8, lRank8); // Outer Forest
    processSpreadsheet(table, ws, "W", 6, 8, 8, lRank8); // Turbine Loop
        processSpreadsheet(table, ws, "B", 6, 19, 8, lRank8); // Treetops
        processSpreadsheet(table, ws, "I", 6, 19, 8, lRank8); // Dark Arsenal
        processSpreadsheet(table, ws, "P", 6, 19, 8, lRank8); // Highway Zero
        processSpreadsheet(table, ws, "W", 6, 19, 8, lRank8); // Bingo Party
    processSpreadsheet(table, ws, "B", 6, 30, 8, lRank8); // Thunder Deck
    processSpreadsheet(table, ws, "I", 4, 30, 3, lRank8); // WR Tally
        for (const auto& entry : table.entries) {
            std::cout << "Entry: " << entry.position << " " << entry.name << " - " << entry.time << std::endl;
            std::cout << entry.racer << std::endl;
        }
        table.clear();
    ws = wb.sheet_by_index(3);  // Alternative Category
    processSpreadsheet(table, ws, "B", 6, 8, 8, lRank8); // Pinball Highway (Springless)
    processSpreadsheet(table, ws, "I", 6, 8, 8, lRank8); // Bingo Party (Springless)
    processSpreadsheet(table, ws, "P", 6, 8, 8, lRank8); // Jump Parage (No Sonic)
        for (const auto& entry : table.entries) {
            std::cout << "Entry: " << entry.position << " " << entry.name << " - " << entry.time << std::endl;
            std::cout << entry.racer << std::endl;
        }
        table.clear();
    ws = wb.sheet_by_index(4); // No-DLC World Records
    processSpreadsheet(table, ws, "B", 6, 8, 1, lRank8); // Icicle Valey
    processSpreadsheet(table, ws, "I", 6, 8, 1, lRank8); // Outer Forest
    processSpreadsheet(table, ws, "P", 6, 8, 1, lRank8); // Turbine Loop
        processSpreadsheet(table, ws, "B", 6, 12, 1, lRank8); // Pinball Highway
        processSpreadsheet(table, ws, "I", 6, 12, 1, lRank8); // Pinball Highway (Springless)
        processSpreadsheet(table, ws, "P", 6, 12, 1, lRank8); // Sewer Scrapes
    processSpreadsheet(table, ws, "B", 6, 17, 2, lRank8); // Sandy Drifts
    processSpreadsheet(table, ws, "I", 6, 17, 1, lRank8); // Deadly Route
    processSpreadsheet(table, ws, "P", 6, 17, 1, lRank8); // Thunder Deck
        for (const auto& entry : table.entries) {
            std::cout << "Entry: " << entry.position << " " << entry.name << " - " << entry.time << std::endl;
            std::cout << entry.racer << std::endl;
        }
        table.clear();
    while (1);
}