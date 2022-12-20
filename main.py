from pathlib import Path
from report.report import Report


def main():
    if __name__ == "__main__":
        path = Path(__file__).parents[0]

        report = Report(path)
        report.start()

        technologies = ("Python", "Golang", "JavaScript")
        tech_row_name = 'title'
        for tech in technologies:
            report.create_tech_worksheet(tech, tech_row_name)

        report.finish()


main()
