#include <iostream>

using namespace std;

const int MAXN = 100;

bool maxPositiveOdd(const int a[], int n, int &value)
{
    bool found = false;

    for (int i = 0; i < n; i++)
    {
        if (a[i] > 0 && a[i] % 2 != 0)
        {
            if (!found || a[i] > value)
            {
                value = a[i];
                found = true;
            }
        }
    }
    return found;
}

bool minPositiveEven(const int a[], int n, int &value)
{
    bool found = false;

    for (int i = 0; i < n; i++)
    {
        if (a[i] > 0 && a[i] % 2 == 0)
        {
            if (!found || a[i] < value)
            {
                value = a[i];
                found = true;
            }
        }
    }
    return found;
}

int removeZeros(int a[], int &n)
{
    int write = 0;

    for (int i = 0; i < n; i++)
    {
        if (a[i] != 0)
            a[write++] = a[i];
    }

    int removed = n - write;
    n = write;
    return removed;
}

void printArray(const int a[], int n)
{
    for (int i = 0; i < n; i++)
    {
        if (i > 0)
            cout << " ";
        cout << a[i];
    }
    cout << endl;
}

int main()
{
    int n;
    int x[MAXN];

    cout << "Лабораторная работа №4" << endl;
    cout << "Циклические процессы с функциями" << endl;
    cout << "Задание 1" << endl << endl;

    cout << "Введите n: ";
    cin >> n;
    if (n <= 0 || n > MAXN)
    {
        cout << "Неверное n." << endl;
        return 0;
    }

    cout << "Введите " << n << " целых чисел: ";
    for (int i = 0; i < n; i++)
        cin >> x[i];

    int oddMax = 0;
    int evenMin = 0;

    cout << endl;
    if (maxPositiveOdd(x, n, oddMax))
        cout << "Макс. положительный нечётный: "
            << oddMax << endl;
    else
        cout << "Положительных нечётных нет." << endl;

    if (minPositiveEven(x, n, evenMin))
        cout << "Мин. положительный чётный: "
            << evenMin << endl;
    else
        cout << "Положительных чётных нет." << endl;

    int removed = removeZeros(x, n);
    cout << "Удалено нулевых элементов: "
        << removed << endl;
    cout << "Массив без нулей: ";
    printArray(x, n);

    return 0;
}
