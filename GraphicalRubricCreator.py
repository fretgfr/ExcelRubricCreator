from tkinter import Tk, Button, Entry, Frame
from tkinter.messagebox import showerror
from tkinter.scrolledtext import ScrolledText
from enum import Enum
import openpyxl

POINTS = False
TOTAL_POINTS = 0

HEADERS = True
STARTING_ROW = 2
HEADER_ROW = 1

FINAL_PERCENTAGE_ROW = 5
FINAL_PERCENTAGE_COLUMN = 7

FINAL_LETTER_ROW = 10
FINAL_LETTER_COLUMN = 7

ASSIGNMENT_NAME_COLUMN = 2
ASSIGNMENT_WEIGHT_COLUMN = 3
ASSIGNMENT_GRADE_COLUMN = 4

LETTER_GRADE_LETTER_COLUMN = 9
LETTER_GRADE_MINIMUM_COLUMN = 10


ASSIGNMENT_WEIGHT_COL_LETTER = chr(ASSIGNMENT_WEIGHT_COLUMN + 64)
ASSIGNMENT_GRADE_COL_LETTER = chr(ASSIGNMENT_GRADE_COLUMN + 64)

class Step(Enum):
    GRADES = 1
    ASSIGNMENTS = 2
    SUBMIT = 3


class RubricCreator(Frame):
    def __init__(self, parent, *args, **kwargs):
        Frame.__init__(self, parent, *args, **kwargs)
        self.parent = parent

        self.step = Step.GRADES

        self.text_area = ScrolledText(self)
        self.entry_area = Entry(self)
        self.submit_button = Button(self, text="Submit", command=lambda: self.submit_option)

        self.text_area.pack()
        self.entry_area.pack(fill='both')
        self.submit_button.pack()

    def submit_option(self):
        text = self.entry_area.get()
    
    def handle_grades(self):
        pass

    def handle_assignments(self):
        pass

    def handle_submit(self):
        pass

    def next(self):
        match self.step:
            case Step.GRADES:
                pass
            case Step.ASSIGNMENTS:
                pass
            case Step.SUBMIT:
                pass
            case _:
                showerror("Something Went Wrong", "Something went wrong. Please report this error on GitHub. ERROR CODE 11")


def main():
    root = Tk()
    RubricCreator(root).pack(side="top", fill='both', expand=True)
    root.mainloop()

if __name__ == "__main__":
    root = Tk()
    rubric_creator = RubricCreator(root)
    rubric_creator.pack(side='top', fill='both', expand=True)
    root.after(10, rubric_creator.start())
    root.mainloop()