from src.dataset import (
    load_data,
    clean_data,
    feature_engineering,
    generate_report,
    create_visualizations,
    generate_profiling_report,
)

LINE = "=" * 60


def main() -> None:
    """
    Execute the complete data analytics workflow.
    """

    # Load dataset
    df = load_data()

    original_rows = len(df)

    # Clean dataset
    df = clean_data(df)

    # Perform feature engineering
    df = feature_engineering(df)

    # Generate PDF report
    generate_report(df)

    # Create charts
    create_visualizations(df)

    # Generate profiling report
    generate_profiling_report(df)

    print(LINE)
    print("PROJECT SUMMARY")
    print(LINE)

    print(f"Original Records : {original_rows}")
    print(f"Final Records    : {len(df)}")
    new_features = len(df.columns) - 12
    print(f"Features Created : {new_features}")
    print("Reports Created  : 1")
    print("Charts Created   : 4")

    print("\nProject completed successfully.")


if __name__ == "__main__":
    main()