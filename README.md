# Export Database Table to Excel

## Requirement
1. Python
2. Pip

## How to run
1. Change the table_names list by query
```sql
SELECT array_to_json(array_agg(table_name))
FROM information_schema.tables
WHERE table_schema = 'public'
  AND table_type = 'BASE TABLE';
```

2. Activate the venv
```bash
source .venv/bin/activate
```

3. Install the requirement
```bash
pip install -r requirements.txt
```

4. Run the program
```bash
python main.py
```