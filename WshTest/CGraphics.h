#pragma once

class CGraphics
{
private:
	int AxisStartX;
	int AxisStartY;

	void DrawXAxis(int x, int y)
	{
		int length = 520;

		//Draw basic line
		MoveToEx(hDC, x, y,  NULL);
		LineTo(hDC, x + length, y);

		//Draw markers
		for (int i = x; i < x + length; i+=50)
		{
			MoveToEx(hDC, i, y - 2, NULL);
			LineTo(hDC, i, y + 3);
		}

		//Draw arrow
		MoveToEx(hDC, x + length, y,  NULL);
		LineTo(hDC, x + length - 5, y - 5);
		LineTo(hDC, x + length - 5, y + 5);
		LineTo(hDC, x + length, y);
	}

	void DrawYAxis(int x, int y)
	{
		int length = 520;

		//Draw basic line
		MoveToEx(hDC, x, y,  NULL);
		LineTo(hDC, x, y - length);

		//Draw markers
		for (int i = y; i > y - length; i-=50)
		{
			MoveToEx(hDC, x - 2, i, NULL);
			LineTo(hDC, x + 3, i);
		}

		//Draw arrow
		MoveToEx(hDC, x, y - length,  NULL);
		LineTo(hDC, x - 5, y - length + 5);
		LineTo(hDC, x + 5, y - length + 5);
		LineTo(hDC, x, y - length);
	}
public:
	HDC hDC;

	CGraphics(void){}
	~CGraphics(void){}

	void DrawAxes(int x, int y)
	{
		AxisStartX = x;
		AxisStartY = y;

		DrawXAxis(AxisStartX, AxisStartY);
		DrawYAxis(AxisStartX, AxisStartY);
	}

	void DrawFast(int buffer[], int length)
	{
		for (int i = 0; i < length; i++)
		{
			SetPixel(hDC, AxisStartX + i, (buffer[i] - AxisStartY) * (-1), NULL);
		}
	}

	void DrawInterp(int buffer[], int length)
	{
		MoveToEx(hDC, AxisStartX, (buffer[0] - AxisStartY) * (-1), NULL);

		for (int i = 1; i < length; i++)
		{
			LineTo(hDC, AxisStartX + i, (buffer[i] - AxisStartY) * (-1));
		}
	}
};
