library(tidyverse)
library(ggthemes)
library(janitor)
library(RDCOMClient)
library(xtable)
library(zoo)

district <- 'Bhopal'
date_format <- '%d/%m/%Y'
subject <- 'Bhopal\'s covid report for today'

df12 <- read_csv('https://api.covid19india.org/csv/latest/raw_data.csv') %>%
  filter(`Detected District` == district) %>%
  rename(date = `Date Announced`) %>%
  count(date, name = 'Hospitalized') %>%
  mutate(Recovered = 0, Deceased = 0,
         date = as.Date(date, format = date_format)) %>%
  pivot_longer(cols = 2:4, names_to = 'status', values_to = 'cases')

df3 <- read_csv('https://api.covid19india.org/csv/latest/raw_data3.csv') %>%
  filter(`Detected District` == district)%>%
  rename(date = `Date Announced`,
         status = `Current Status`,
         cases = `Num Cases`) %>%
  mutate(date = as.Date(date, format = date_format)) %>%
  select(date, status, cases) %>% 
  pivot_wider(names_from = status, values_from = cases) %>% 
  replace_na(list(Hospitalized = 0, Recovered = 0, Deceased = 0)) %>%
  pivot_longer(cols = 2:4, names_to = 'status', values_to = 'cases') 

df4 <- read_csv('https://api.covid19india.org/csv/latest/raw_data4.csv') %>%
  filter(`Detected District` == district)%>%
  rename(date = `Date Announced`,
         status = `Current Status`,
         cases = `Num Cases`) %>%
  mutate(date = as.Date(date, format = date_format)) %>%
  select(date, status, cases) %>% 
  pivot_wider(names_from = status, values_from = cases) %>% 
  replace_na(list(Hospitalized = 0, Recovered = 0, Deceased = 0)) %>%
  pivot_longer(cols = 2:4, names_to = 'status', values_to = 'cases') 

df5 <- read_csv('https://api.covid19india.org/csv/latest/raw_data5.csv') %>%
  filter(`Detected District` == district)%>%
  rename(date = `Date Announced`,
         status = `Current Status`,
         cases = `Num Cases`) %>%
  mutate(date = as.Date(date, format = date_format)) %>%
  select(date, status, cases) %>% 
  pivot_wider(names_from = status, values_from = cases) %>% 
  replace_na(list(Hospitalized = 0, Recovered = 0, Deceased = 0)) %>%
  pivot_longer(cols = 2:4, names_to = 'status', values_to = 'cases') 

df <- rbind.data.frame(df12, df3, df4, df5)

df_wide <- df %>%
  arrange(date) %>%
  pivot_wider(names_from = status, values_from = cases) %>%
  rename(Date = date,
         `New Cases` = Hospitalized,
         Recoveries = Recovered,
         Deaths = Deceased) %>%
  mutate(`Total Cases` = cumsum(`New Cases`),
         `Total Recoveries` = cumsum(Recoveries),
         `Total Deaths` = cumsum(Deaths), 
         `Active Cases` = `Total Cases` - `Total Deaths` - `Total Recoveries`) %>%
  select(-`Total Cases`, -`Total Recoveries`, -`Total Deaths`) %>%
  adorn_totals(where = 'row') 

df_wide[nrow(df_wide), ncol(df_wide)] <- df_wide[nrow(df_wide)-1, ncol(df_wide)]

df_total <- df %>%
  arrange(date) %>%
  pivot_wider(names_from = status, values_from = cases) %>%
  rename(Date = date,
         `New Cases` = Hospitalized,
         Recoveries = Recovered,
         Deaths = Deceased) %>%
  mutate(`Total Cases` = cumsum(`New Cases`),
         `Total Recoveries` = cumsum(Recoveries),
         `Total Deaths` = cumsum(Deaths),
         `Active Cases` = `Total Cases` - `Total Deaths` - `Total Recoveries`) %>%
  select(-`New Cases`, -Recoveries, -Deaths) %>%
  pivot_longer(cols = 2:5, names_to = 'status', values_to = 'cases')

df_plot <- df_wide %>%
  mutate(`Total Cases` = cumsum(`New Cases`),
         `Total Recoveries` = cumsum(Recoveries),
         `Total Deaths` = cumsum(Deaths),
         `Active Cases` = `Total Cases` - `Total Deaths` - `Total Recoveries`) %>%
  select(-`New Cases`, -Recoveries, -Deaths) %>%
  mutate(`Total Cases MA` = rollmean(`Total Cases`, k=5, fill = NA),
         `Total Recoveries MA` = rollmean(`Total Recoveries`, k=5, fill = NA),
         `Total Deaths MA` = rollmean(`Total Deaths`, k=5, fill = NA))

df_wide <- df_wide %>%
  mutate_all(as.character)

html_attachment <- print(xtable(df_wide, align = rep('c', times = 6)),
                         type="html", print.results=FALSE)
html_attachment <- paste0(paste0("<html>", html_attachment, "</html>"))

daily_plot <- ggplot(df) +
  aes(x = date, y = cases, color = status) +
  geom_point(alpha = 0.5, size = 2) +
  stat_smooth(size = 2, se = FALSE) +
  scale_color_manual(values = c('red', 'orange', 'darkblue')) +
  labs(title = 'Daily Cases (Bhopal)') +
  theme_wsj()

print(daily_plot)

ggsave(filename = 'daily-plot.PNG',
       plot = daily_plot,
       dpi = 1000)

cumulative_plot <- ggplot(df_total) +
  aes(x = Date, y = cases, color = status) +
  geom_point(alpha = 0.5, size = 2) +
  geom_point(color = 'black', shape = 21, size = 2) +
  stat_smooth(size = 1.5, se = FALSE) +
  scale_color_manual(values = c('darkblue', 'darkorange', 'red', 'green4')) +
  labs(title = 'Bhopal\'s Covid-19 Trajectory',
       caption = 'Tanmay Dixit') +
  theme_wsj() +
  theme(plot.title = element_text(hjust = 0.5),
        plot.caption = element_text(size = 10, face = 'bold'),
        legend.position = 'top') 

print(cumulative_plot)

ggsave(filename = 'cumulative-plot.PNG', 
       plot = cumulative_plot,
       width = 12, height = 6,
       dpi = 1000)

Outlook <- COMCreate('Outlook.Application')

Email = Outlook$CreateItem(0)
Email[['to']] = 'your_email@domain.com'
Email[['subject']] = subject 
Email[['htmlbody']] = html_attachment
Email[["attachments"]]$Add('cumulative-plot.PNG')
Email$Send()
