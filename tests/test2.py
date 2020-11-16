"GDP per capita. Example from Plotly docs."

import plotly.express as px
df = px.data.gapminder()

fig = px.scatter(df.query("year==2007"), x="gdpPercap", y="lifeExp",
	         size="pop", color="continent",
                 hover_name="country", log_x=True, size_max=60)
fig.write_image("test2.svg", width=800, height=800)
# fig.show()
