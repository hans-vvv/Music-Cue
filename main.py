import logging

import ttkbootstrap as ttk
import application


if __name__ == '__main__':
    logging.basicConfig(filename='logging.txt', filemode='w',
                        format='%(asctime)s - %(levelname)s: %(message)s', level=logging.ERROR)
    app = ttk.Window(
            title="MusicCue",
            themename="superhero",
            size=(1500, 800),
            resizable=(True, True),
        )
    application.MusicCueGui(app)
    app.mainloop()
