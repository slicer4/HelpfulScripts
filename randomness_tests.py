import numpy as np
import scipy.stats as stats

class NumPyRandomnessValidator:
    def __init__(self, sample_size=10000):
        self.sample_size = sample_size
        self.rng = np.random.default_rng()
        self.samples = None

    def generate_samples(self):
        """Generate uniform random samples"""
        self.samples = self.rng.uniform(0, 1, self.sample_size)
        return self.samples

    def kolmogorov_smirnov_test(self):
        """
        Kolmogorov-Smirnov Test
        Compares sample distribution to uniform distribution
        """
        if self.samples is None:
            self.generate_samples()
        
        # Perform KS test against uniform distribution
        ks_statistic, p_value = stats.kstest(self.samples, 'uniform')
        
        print("\nKolmogorov-Smirnov Test:")
        print(f"KS Statistic: {ks_statistic:.4f}")
        print(f"P-value: {p_value:.4f}")
        print("Interpretation: ")
        if p_value > 0.05:
            print("✓ Samples appear to be uniformly distributed")
        else:
            print("⚠ Significant deviation from uniform distribution")
        
        return {
            'statistic': ks_statistic, 
            'p_value': p_value,
            'is_uniform': p_value > 0.05
        }

    def chi_squared_test(self, num_bins=10):
        """
        Chi-Squared Test
        Checks if observed frequencies match expected frequencies
        """
        if self.samples is None:
            self.generate_samples()
        
        # Create histogram
        counts, _ = np.histogram(self.samples, bins=num_bins)
        
        # Expected count per bin (uniform distribution)
        expected_count = self.sample_size / num_bins
        expected_frequencies = np.full(num_bins, expected_count)
        
        # Chi-squared calculation
        chi2_statistic, p_value = stats.chisquare(counts, f_exp=expected_frequencies)
        
        print("\nChi-Squared Test:")
        print(f"Bin Counts: {counts}")
        print(f"Chi-squared Statistic: {chi2_statistic:.4f}")
        print(f"P-value: {p_value:.4f}")
        print("Interpretation: ")
        if p_value > 0.05:
            print("✓ Frequencies consistent with uniform distribution")
        else:
            print("⚠ Significant deviation in bin frequencies")
        
        return {
            'statistic': chi2_statistic, 
            'p_value': p_value,
            'is_uniform': p_value > 0.05,
            'bin_counts': counts
        }

    def runs_test(self):
        """
        Runs Test
        Checks for randomness by analyzing sequence of above/below median runs
        """
        if self.samples is None:
            self.generate_samples()
        
        # Calculate median
        median = np.median(self.samples)
        
        # Create binary sequence (above/below median)
        binary_seq = (self.samples > median).astype(int)
        
        # Count runs
        runs = np.diff(binary_seq).nonzero()[0].size + 1
        
        # Calculate expected runs and standard deviation
        n1 = np.sum(binary_seq)
        n2 = self.sample_size - n1
        
        # Expected number of runs
        expected_runs = ((2 * n1 * n2) / (n1 + n2)) + 1
        
        # Standard deviation of runs
        std_runs = np.sqrt((2 * n1 * n2 * (2 * n1 * n2 - n1
