'''
This program is used to create graphs from the responses generated from a google form
Documentation limitations (hardcodes): questions range from 1-54, responses range from 1,5.
Author: Cooper Goddard
Date (init): 2022-11-01 
'''
import matplotlib.pyplot as plt
import csv

NUMBER_OF_QUESTIONS = 54
NUMBER_OF_RESPONSES = 35
MIN_ANSWER = 1
MAX_ANSWER = 5


def get_csv_data(filename: str) -> list[list[int]] and list[str] and list[str]:
    '''
    Takes a csv file full of data (with a header) and breaks it into 
    a grid array which it returns.
    '''
    with open(filename, "r") as file:
        rows = list(csv.reader(file, delimiter=','))
        headers = rows.pop(0)

        # Populate 2D Array
        grid = []
        comments = []
        for row in rows:
            row.pop(0)  # remove timestamp
            comments.append(row.pop(-1))  # remove comment
            grid.append(row)

        # Make a new grid, one row for each question
        new_grid = []
        count = 0
        while count < NUMBER_OF_QUESTIONS-1:
            new_grid.append([])
            count += 1

        # swap rows and columns, adding to the new grid
        for row in range(len(grid)):
            for value in range(len(grid[row])):
                new_grid[value].append(int(grid[row][value]))

    return new_grid, headers, comments


def sort_data(data: list[int]) -> list[int]:
    '''
    This function takes a list of numbers and totals the ammount of each number
    For example: [1,1,1,2,3,4,3] would return [3,1,2,1,0]
    currently assumes that MIN_ANSWER > 0
    '''
    result = [0] * MAX_ANSWER
    for value in data:
        # add count for the value to proper index
        result[value-1] += 1

    return result


def create_graph(header: str, data: list[int]):
    '''
    This function creates a graph based on the column provided (question num)
    using the 2d Array 'grid'. It adds the header and returns an Image
    '''
    fig, ax = plt.subplots()
    bar_labels = ['1', '2', '3', '4', '5']
    bar_colors = ['tab:green', 'tab:blue', 'tab:red', 'tab:orange', 'tab:pink']

    ax.bar(bar_labels, data, label=bar_labels, color=bar_colors)
    ax.set_ylim(0, NUMBER_OF_RESPONSES)
    ax.bar_label(ax.containers[0], label_type='edge')
    ax.set_ylabel('Number of Responses')
    ax.set_title(header)

    plt.show()


def make_document(grid: list[list[int]], headers: list[str]) -> None:
    '''
    This function uses the data to build graphs and insert them into a
    word document (.docx)
    '''
    # create document to be populated

    # For each question, create a graph and upload to the document
    for question in range(len(grid)):
        create_graph(headers[question], sort_data(grid[question]))


def main() -> None:
    # Get Data From CSV
    grid, headers, comments = get_csv_data('responses.csv')
    headers.pop(0)  # remove timestamp
    headers.pop(-1)  # remove comment header
    # Create Graphs for each question and upload to document
    make_document(grid, headers)


if __name__ == "__main__":
    main()
