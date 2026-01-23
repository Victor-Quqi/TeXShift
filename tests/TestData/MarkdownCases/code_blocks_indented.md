# Indented code blocks (4 spaces)

目标：验证 4 空格缩进的代码块被正确识别为 code block，并且与普通段落分隔正常。

Basic indented block:

    using System;
    class Program
    {
        static void Main()
        {
            Console.WriteLine("Hello from indented code.");
        }
    }

Indented block with blank lines:

    def fib(n):
        if n <= 1:
            return n

        return fib(n - 1) + fib(n - 2)

Text after the code block should be normal paragraph text.

