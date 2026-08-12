#include <fstream>
#include <iomanip>
#include <iostream>
#include <sstream>
#include <string>

int main(int argc, char* argv[]) {
    if (argc != 2) {
        std::cerr << "Usage: wordcount <filename>\n";
        return 1;
    }

    std::ifstream input(argv[1]);
    if (!input.is_open()) {
        std::cerr << "Error: could not open file \"" << argv[1] << "\"\n";
        return 1;
    }

    std::string word;
    std::string line;
    int lineCount = 0;
    int wordCount = 0;
    int charCount = 0;

    while (std::getline(input, line)) {
        ++lineCount;
        charCount += static_cast<int>(line.length()) + 1;

        std::istringstream stream(line);
        while (stream >> word) {
            ++wordCount;
        }
    }

    std::cout << std::left << std::setw(10) << "Lines: " << lineCount << "\n"
              << std::setw(10) << "Words: " << wordCount << "\n"
              << std::setw(10) << "Chars: " << charCount << "\n";

    return 0;
}
