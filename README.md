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

const blob = new Blob([converter.data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = 'data.xlsx';
a.click();
URL.revokeObjectURL(url);
```

