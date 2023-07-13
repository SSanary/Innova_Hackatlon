import xlwings as xw
import random

beklenen=[]
metrics=[]
data_metrics=[]

population_size = 100
population = []
fitnes_of_population = []
best_individual = []
Best_Score = 1500

path = "C:\\Users\\elifn\\OneDrive\\Belgeler\\GitHub\\Innova_Future_Istanbul_Hackatlon\\Innova_Hackatlon\\assets\\target-file.xlsx"

def readData():
    # Getting the data from beklenen sheet
    ws = xw.Book(path).sheets['beklenen']
    for letter in range(66, 90):
        beklenen.append([])
        for i in range(2,14):
            beklenen[letter-66].append(ws.range(str(chr(letter)) + str(i)).value)
        
    # Getting the data from metrics sheet
    ws = xw.Book(path).sheets['metrics']
    for letter in range(67, 91):
        metrics.append([])
        for i in range(2,14):
            metrics[letter-67].append(ws.range(str(chr(letter)) + str(i)).value)

    # Getting the data from data_metrics sheet
    ws = xw.Book(path).sheets['data_metrics']
    for letter in range(66, 90):
        data_metrics.append([])
        for i in range(2,13):
            data_metrics[letter-66].append(ws.range(str(chr(letter)) + str(i)).value)

def fitness_function(individual):
    fitness = 0
    global Best_Score
    selection_count_for_outputs = []
    for i in range(len(individual)):
        selection_count_for_outputs.append(0)
        for j in range(len(individual[i])):
            if individual[i][j] == 1:
                selection_count_for_outputs[i] += 1
    for i in range(len(data_metrics)):
        for j in range(len(data_metrics[i])):
            fitness += fitness + ((data_metrics[i][j]) * selection_count_for_outputs[j])/5
    if fitness/23 < Best_Score:
        Best_Score = fitness
        best_individual = individual
    return fitness/24

def createInitialPopulation():
    for i in range(population_size):
        population.append([])
        for j in range(24):
            population[i].append([])
            for k in range(12):
                population[i][j].append(0)
            population[i][j][random.randint(0,11)] = 1
    for i in range(population_size):
        fitnes_of_population.append(fitness_function(population[i]))

def crossover():
    child = []
    child_fitness = 0
    selection_1 = random.randint(0, population_size-1)
    selection_2 = random.randint(0, population_size-1)
    while selection_1 != selection_2:
        selection_2 = random.randint(0, population_size-1)
    crossover_point = random.randint(0, 23)
    for i in range(len(population[selection_1][0])):
        child.append([])
        for j in range(crossover_point):
            child[i].append(population[selection_2][j][i])
        for j in range(crossover_point, len(population[selection_1])):
            child[i].append(population[selection_1][j][i])
    child_fitness = fitness_function(child)
    #TODO: local search
    if fitnes_of_population[selection_1] > fitnes_of_population[selection_2]:
        if child_fitness < fitnes_of_population[selection_1]:
            population[selection_1] = child
            fitnes_of_population[selection_1] = child_fitness
        elif child_fitness < fitnes_of_population[selection_2]:
            population[selection_2] = child
            fitnes_of_population[selection_2] = child_fitness
    else:
        if child_fitness < fitnes_of_population[selection_2]:
            population[selection_2] = child
            fitnes_of_population[selection_2] = child_fitness
        elif child_fitness < fitnes_of_population[selection_1]:
            population[selection_1] = child
            fitnes_of_population[selection_1] = child_fitness
 
def geneticAlgorithm():
    createInitialPopulation()
    for i in range(100):
        crossover()


def main():
    readData()
    geneticAlgorithm()
    print(best_individual)
    
main()