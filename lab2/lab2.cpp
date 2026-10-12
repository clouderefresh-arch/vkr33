#include <iostream>
#include <cmath>
#include <iomanip>

using namespace std;

void triangle()
{
    double a, b, c;

    cout << "Введите стороны a, b, c: ";
    cin >> a >> b >> c;

    if (a <= 0 || b <= 0 || c <= 0
        || a + b <= c || a + c <= b || b + c <= a)
    {
        cout << "Треугольник построить нельзя." << endl;
        return;
    }

    double p = (a + b + c) / 2.0;
    double s = sqrt(p * (p - a) * (p - b) * (p - c));

    cout << fixed << setprecision(6);
    cout << "Треугольник существует." << endl;
    cout << "Полупериметр p = " << p << endl;
    cout << "Площадь S = " << s << endl;
}

void month()
{
    int n;

    cout << "Введите номер месяца (1-12): ";
    cin >> n;

    switch (n)
    {
    case 1:
        cout << "январь, зима" << endl;
        break;
    case 2:
        cout << "февраль, зима" << endl;
        break;
    case 3:
        cout << "март, весна" << endl;
        break;
    case 4:
        cout << "апрель, весна" << endl;
        break;
    case 5:
        cout << "май, весна" << endl;
        break;
    case 6:
        cout << "июнь, лето" << endl;
        break;
    case 7:
        cout << "июль, лето" << endl;
        break;
    case 8:
        cout << "август, лето" << endl;
        break;
    case 9:
        cout << "сентябрь, осень" << endl;
        break;
    case 10:
        cout << "октябрь, осень" << endl;
        break;
    case 11:
        cout << "ноябрь, осень" << endl;
        break;
    case 12:
        cout << "декабрь, зима" << endl;
        break;
    default:
        cout << "Неверный номер месяца." << endl;
        break;
    }
}

int main()
{
    int choice;

    cout << "Лабораторная работа №2" << endl;
    cout << "Ветвящиеся вычислительные процессы" << endl;
    cout << "Задания 1 и 2" << endl;

    do
    {
        cout << endl;
        cout << "Задание (1 - треугольник, "
            << "2 - месяц, 0 - выход): ";
        cin >> choice;

        if (choice == 1)
            triangle();
        else if (choice == 2)
            month();
        else if (choice != 0)
            cout << "Неверный пункт меню." << endl;
    }
    while (choice != 0);

    return 0;
}
