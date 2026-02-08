# Fenced code blocks (languages + fence variants)

## JavaScript
```javascript
const greeting = "Hello, World!";
const items = [1, 2, 3].map(x => x * 2);
console.log({ greeting, items });
```

## Python
```python
from dataclasses import dataclass

@dataclass
class Point:
    x: int
    y: int

print(Point(1, 2))
```

## C#
```csharp
using System;

public static class Demo
{
    public static void Main()
    {
        Console.WriteLine("Hello from C#");
    }
}
```

## SQL
```sql
SELECT u.id, u.username, COUNT(o.id) AS order_count
FROM users u
LEFT JOIN orders o ON u.id = o.user_id
GROUP BY u.id, u.username
ORDER BY order_count DESC;
```

## HTML
```html
<!doctype html>
<html lang="zh-CN">
  <head><meta charset="utf-8"><title>Demo</title></head>
  <body><p>Hello</p></body>
</html>
```

## JSON
```json
{
  "name": "TeXShift",
  "enabled": true,
  "features": ["markdown", "latex"]
}
```

## YAML
```yaml
name: TeXShift
enabled: true
features:
  - markdown
  - latex
```

## Fence variant (`~~~`) + backticks inside
~~~text
This block uses tildes, so it can safely contain backticks:

```json
{"nested": true}
```
~~~

