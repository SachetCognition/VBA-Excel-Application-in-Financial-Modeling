.PHONY: install build test extract clean lint run

install:
	pip install -r requirements.txt

build: install extract

extract:
	python3 scripts/extract_vba.py

test:
	python3 -m pytest tests/ -v

lint:
	python3 scripts/lint_vba.py

run:
	python3 scripts/extract_vba.py --verbose

clean:
	rm -rf build/vba_modules/ __pycache__ tests/__pycache__ .pytest_cache
