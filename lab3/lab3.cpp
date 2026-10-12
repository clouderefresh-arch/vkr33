#include <iostream>
#include <cmath>
#include <iomanip>

using namespace std;

void figure()
{
    for (int row = 0; row < 7; row++)
    {
        int stars;
        if (row < 4)
            stars = 1 + 2 * row;
        else
            stars = 1 + 2 * (6 - row);

        int spaces = (7 - stars) / 2;
        for (int i = 0; i < spaces; i++)
            cout << " ";
        for (int i = 0; i < stars; i++)
            cout << "*";
        cout << endl;
    }
}

void series()
{
    int n;
    cout << "Введите n: ";
    cin >> n;

    if (n <= 0)
    {
        cout << "n должно быть больше 0." << endl;
        return;
    }

    const double pi = acos(-1.0);
    double sum = 0.0;
    double sign = -1.0;
    double denom = 2.0;

    for (int k = 1; k <= n; k++)
    {
        sum += sign * sin(pi / denom);
        sign = -sign;
        denom *= 2.0;
    }

    cout << fixed << setprecision(6);
    cout << "S = " << sum << endl;
}

int main()
{
    int choice;

    cout << "Лабораторная работа №3" << endl;
    cout << "Циклические вычислительные процессы" << endl;
    cout << "Задания 1 и 6" << endl;

    do
    {
        cout << endl;
        cout << "Задание (1 - фигура, 6 - сумма S, "
            << "0 - выход): ";
        cin >> choice;

        if (choice == 1)
            figure();
        else if (choice == 6)
            series();
        else if (choice != 0)
            cout << "Неверный пункт меню." << endl;
    }
    while (choice != 0);

    return 0;
}
