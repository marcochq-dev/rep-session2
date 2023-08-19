using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using OfficeOpenXml;

namespace TSPSimulatedAnnealing
{
    class Program
    {
        static Random random = new Random();

        static double CalculateTotalDistance(List<int> tour, double[,] distances)
        {
            double totalDistance = 0;
            for (int i = 0; i < tour.Count - 1; i++)
            {
                totalDistance += distances[tour[i], tour[i + 1]];
            }
            totalDistance += distances[tour[tour.Count - 1], tour[0]]; // inicio
            return totalDistance;
        }

        static List<int> SwapCities(List<int> tour, int index1, int index2)
        {
            List<int> newTour = new List<int>(tour);
            int temp = newTour[index1];
            newTour[index1] = newTour[index2];
            newTour[index2] = temp;
            return newTour;
        }

        static void Main(string[] args)
        {
            var filePath = "archivo.xlsx";
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            using (var package = new ExcelPackage())
            {
                var cityDataSets = CityData.GetCityDataSets();
                int countfor = 0;
                foreach (var dataSet in cityDataSets)
                {
                    countfor++;

                    // Agregar una hoja al archivo Excel
                    var worksheet = package.Workbook.Worksheets.Add("Datos " + countfor);

                    // Agregar datos a la hoja
                    worksheet.Cells["A1"].Value = "Ciudad";
                    worksheet.Cells["B1"].Value = "Temperatura";
                    worksheet.Cells["C1"].Value = "Enfriamiento";
                    worksheet.Cells["D1"].Value = "Ciclos";
                    worksheet.Cells["E1"].Value = "Aceptación";
                    worksheet.Cells["F1"].Value = "Mejor Ruta";
                    worksheet.Cells["G1"].Value = "Mejor Distancia";
                    worksheet.Cells["H1"].Value = "Tiempo de Proceso (segundos)";

                    int numCities = dataSet.NumCities;
                    double[,] distances = dataSet.Distances;

                    // Ejecutar el algoritmo 10 veces
                    for (int run = 1; run <= 1; run++)
                    {
                        List<int> currentTour = Enumerable.Range(0, numCities).ToList();
                        currentTour = currentTour.OrderBy(x => random.Next()).ToList();

                        double currentDistance = CalculateTotalDistance(currentTour, distances);

                        double initialTemperature = 1000;
                        double coolingRate = 0.000000003;

                        List<int> bestTour = new List<int>(currentTour);
                        double bestDistance = currentDistance;

                        double temperature = initialTemperature;
                        int valcount = 0;
                        int valcountProb = 0;
                        Stopwatch stopwatch = new Stopwatch();
                        stopwatch.Start(); 
                        while (temperature > 1)
                        {
                            valcount++;
                            int randomIndex1 = random.Next(0, numCities);
                            int randomIndex2 = random.Next(0, numCities);

                            List<int> newTour = SwapCities(currentTour, randomIndex1, randomIndex2);
                            double newDistance = CalculateTotalDistance(newTour, distances);

                            double acceptanceProbability = Math.Exp((currentDistance - newDistance) / temperature);

                            if (newDistance < currentDistance || random.NextDouble() < acceptanceProbability)
                            {
                                valcountProb ++;
                                currentTour = newTour;
                                currentDistance = newDistance;

                                if (currentDistance < bestDistance)
                                {
                                    
                                    bestTour = new List<int>(currentTour);
                                    bestDistance = currentDistance;
                                }
                            }

                            temperature *= 1 - coolingRate;
                        }

                        stopwatch.Stop(); // Detiene el cronómetro

                        bestTour.Add(bestTour[0]);
                        Console.WriteLine(" contador: " + valcount);
                        Console.WriteLine(" probabilidad: " + valcountProb);
                        Console.WriteLine("Best Tour: " + string.Join(" -> ", bestTour));
                        Console.WriteLine("Best Distance: " + bestDistance);

                        // Almacenar el tiempo de proceso en segundos en una variable
                        var tiempoProceso = stopwatch.Elapsed.TotalSeconds;
                        Console.WriteLine("Tiempo Proceso: " + tiempoProceso);

                        // Ejemplo de datos
                        var ciudad = numCities;
                        var temperatura = initialTemperature;
                        var enfriamiento = coolingRate;
                        var ciclos = valcount;
                        var aceptacion = valcountProb;
                        var mejorRuta = string.Join(" -> ", bestTour);
                        var mejorDistancia = bestDistance;

                        // Agregar datos al archivo Excel a partir de la fila 3
                        int row = run + 2;
                        worksheet.Cells["A" + row].Value = ciudad;
                        worksheet.Cells["B" + row].Value = temperatura;
                        worksheet.Cells["C" + row].Value = enfriamiento;
                        worksheet.Cells["D" + row].Value = ciclos;
                        worksheet.Cells["E" + row].Value = aceptacion;
                        worksheet.Cells["F" + row].Value = mejorRuta;
                        worksheet.Cells["G" + row].Value = mejorDistancia;
                        worksheet.Cells["H" + row].Value = tiempoProceso;
                    }
                }
                // Guardar el archivo Excel en la ubicación especificada
                File.WriteAllBytes(filePath, package.GetAsByteArray());
            }
            Console.WriteLine("Archivo Excel guardado en: " + filePath);
        }
    }
}
