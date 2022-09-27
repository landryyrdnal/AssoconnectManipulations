from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.image import Image

class BaseDeDonneAssoConnectApp(App):
    def build(self):
        box = BoxLayout(orientation="vertical")
        label = Label(text="bonjour")
        button = Button(text="bouton de test",
                      background_color=(0,0,1,1),
                      font_size = 120)
        picture = Image(source="icon_nini.png")
        box.add_widget(button)
        box.add_widget(label)
        box.add_widget(picture)
        return box
    pass

if __name__ == "__main__":
    BaseDeDonneAssoConnectApp().run()