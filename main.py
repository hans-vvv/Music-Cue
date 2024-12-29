import ttkbootstrap as ttk
import application


if __name__ == '__main__':

    app = ttk.Window(
            title="MusicCue",
            themename="superhero",
            size=(1500, 800),
            resizable=(True, True),
        )
    application.MusicCueGui(app)
    app.mainloop()
