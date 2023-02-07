import csv
import logging


def create_csv(rows, file_name, header, file_path, encoding, delimiter):
    full_path = f'{file_path}/{file_name}.csv'
    logging.info(f"Creating... {full_path}")
    with open(full_path, 'w', encoding=encoding, newline='') as file:
        writer = csv.writer(file, delimiter=delimiter)
        writer.writerow(header)
        writer.writerows(rows)
    logging.info(f"Created: {full_path}")


def read_csv(file_path, delimiter, encoding):
    csv_rows = []
    logging.info(f'Reading... {file_path}')
    with open(file_path, encoding=encoding) as csv_file:
        reader = csv.reader(csv_file, delimiter=delimiter)
        for row in reader:
            csv_rows.append(row)
    return csv_rows


def write_csv(rows, file_path, encoding):
    with open(file_path, 'a', newline='', encoding=encoding) as file:
        writer = csv.writer(file, delimiter=',')
        writer.writerow(rows)
