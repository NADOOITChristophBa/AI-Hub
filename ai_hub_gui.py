import toga
from toga.style import Pack
from toga.style.pack import COLUMN, ROW


# Define the submit_command function which will be called when the button is pressed
def submit_command(widget):
    # Placeholder for command submission logic
    print("Command submitted")


def build(app):
    # Create a main window with a title
    main_window = toga.MainWindow(title="AI Hub")

    # Create a button
    button = toga.Button("Submit", on_press=submit_command, style=Pack(padding=5))

    # Create text input box
    text_input = toga.TextInput(
        placeholder="Type your command here", style=Pack(flex=1, padding=5)
    )

    # Create an output box - multi-line text input that is read-only
    output_box = toga.MultilineTextInput(
        readonly=True, style=Pack(flex=5)
    )  # flex is higher to take up remaining space

    # Create a box to hold the input and button
    input_box = toga.Box(
        children=[text_input, button], style=Pack(direction=ROW, padding=5)
    )

    # Create a box to hold everything, with the output box taking most of the space
    main_box = toga.Box(children=[output_box, input_box], style=Pack(direction=COLUMN))

    # Add the main box to the main window
    main_window.content = main_box

    return main_box


# Create the Toga application
app = toga.App("AI Hub", "org.beeware.helloworld", startup=build)

# Run the Toga application
app.main_loop()
