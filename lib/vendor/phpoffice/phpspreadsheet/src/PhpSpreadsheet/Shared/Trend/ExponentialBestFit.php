<?php

namespace PhpOffice\PhpSpreadsheet\Shared\Trend;

class ExponentialBestFit extends BestFit
{
    /**
     * Algorithm type to use for best-fit
     * (Name of this Trend class).
     *
     * @var string
     */
    protected $bestFitType = 'exponential';

    /**
     * Return the Y-Base for a specified value of X.
     *
     * @param float $xValue X-Base
     *
     * @return float Y-Base
     */
    public function getValueOfYForX($xValue)
    {
        return $this->getIntersect() * $this->getSlope() ** ($xValue - $this->xOffset);
    }

    /**
     * Return the X-Base for a specified value of Y.
     *
     * @param float $yValue Y-Base
     *
     * @return float X-Base
     */
    public function getValueOfXForY($yValue)
    {
        return log(($yValue + $this->yOffset) / $this->getIntersect()) / log($this->getSlope());
    }

    /**
     * Return the Equation of the best-fit line.
     *
     * @param int $dp Number of places of decimal precision to display
     *
     * @return string
     */
    public function getEquation($dp = 0)
    {
        $slope = $this->getSlope($dp);
        $intersect = $this->getIntersect($dp);

        return 'Y = ' . $intersect . ' * ' . $slope . '^X';
    }

    /**
     * Return the Slope of the line.
     *
     * @param int $dp Number of places of decimal precision to display
     *
     * @return float
     */
    public function getSlope($dp = 0)
    {
        if ($dp != 0) {
            return round(exp($this->slope), $dp);
        }

        return exp($this->slope);
    }

    /**
     * Return the Base of X where it intersects Y = 0.
     *
     * @param int $dp Number of places of decimal precision to display
     *
     * @return float
     */
    public function getIntersect($dp = 0)
    {
        if ($dp != 0) {
            return round(exp($this->intersect), $dp);
        }

        return exp($this->intersect);
    }

    /**
     * Execute the regression and calculate the goodness of fit for a set of X and Y data values.
     *
     * @param float[] $yValues The set of Y-values for this regression
     * @param float[] $xValues The set of X-values for this regression
     */
    private function exponentialRegression(array $yValues, array $xValues, bool $const): void
    {
        $adjustedYValues = array_map(
            function ($value) {
                return ($value < 0.0) ? 0 - log(abs($value)) : log($value);
            },
            $yValues
        );

        $this->leastSquareFit($adjustedYValues, $xValues, $const);
    }

    /**
     * Define the regression and calculate the goodness of fit for a set of X and Y data values.
     *
     * @param float[] $yValues The set of Y-values for this regression
     * @param float[] $xValues The set of X-values for this regression
     * @param bool $const
     */
    public function __construct($yValues, $xValues = [], $const = true)
    {
        parent::__construct($yValues, $xValues);

        if (!$this->error) {
            $this->exponentialRegression($yValues, $xValues, (bool) $const);
        }
    }
}
