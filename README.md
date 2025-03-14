# simple-xlsx-writer (WIP)

## Usage 

```ts
import { Converter } from 'simple-xlsx-writer';

const headers = [
  { label: "Name", key: "firstName" },
  { label: "Surname", key: "lastName" },
  { label: "Age", key: "age" }
]

const data = [
  { firstName: "Alice", lastName: "Anderson", age: 34 },
  { firstName: "Bob", lastName: "Brown", age: 29 },
  { firstName: "Charlie", lastName: "Clark", age: 42 }
]

const converter = new Converter();
converter.append({ headers, data, sheetName: 'People' });
converter.save("people.xlsx");
```

This will create an xlsx file with this structure:

| Name | Surname | Age | Country | City     | Address    | Phone       | Email              |
|------|---------|-----|---------|----------|------------|-------------|--------------------|
| John | Doe     | 25  | USA     | New York | 123 Street | 123-456-789 | example@gmail.com  |
| John | Doe     | 25  | USA     | New York | 123 Street | 123-456-789 | example@gmail.com  |
| John | Doe     | 25  | USA     | New York | 123 Street | 123-456-789 | example@gmail.com  |

## API

| Method       | Return Type   | Details                                |
|--------------|---------------|----------------------------------------|
| append       | void          | Synchronously appends a new worksheet. |
| append_async | Promise<void> | Asynchronously appends a new worksheet.|
| save         | void          | Saves the xlsx to file.                |

