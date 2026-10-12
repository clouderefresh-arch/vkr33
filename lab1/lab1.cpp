#include <iostream>
#include <cmath>
#include <iomanip>

using namespace std;

int main()
{
    const double pi = acos(-1.0);
    double alpha;

    cout << "Лабораторная работа №1" << endl;
    cout << "Линейные вычислительные процессы" << endl;
    cout << "Варианты 1 и 2" << endl << endl;
    cout << "Введите alpha (радианы): ";
    cin >> alpha;

    double y1 = 2.0 * pow(sin(3.0 * pi - 2.0 * alpha), 2.0)
        * pow(cos(5.0 * pi + 2.0 * alpha), 2.0);
    double z1 = 1.0 / 4.0
        - (1.0 / 4.0) * sin(5.0 / 2.0 * pi - 8.0 * alpha);

    double y2 = cos(alpha) + sin(alpha)
        + cos(3.0 * alpha) + sin(3.0 * alpha);
    double z2 = 2.0 * sqrt(2.0) * cos(alpha)
        * sin(pi / 4.0 + 2.0 * alpha);

    cout << endl << fixed << setprecision(10);
    cout << "Вариант 1" << endl;
    cout << "y     = " << y1 << endl;
    cout << "z     = " << z1 << endl;
    cout << "y - z = " << scientific << setprecision(3)
        << (y1 - z1) << endl;

    cout << endl << fixed << setprecision(10);
    cout << "Вариант 2" << endl;
    cout << "y     = " << y2 << endl;
    cout << "z     = " << z2 << endl;
    cout << "y - z = " << scientific << setprecision(3)
        << (y2 - z2) << endl;

    return 0;
}
